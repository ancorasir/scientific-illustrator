"""Local recursive connector audit and scoped, optimistic PPTX patching.

This module never contacts a service or saves through a presentation renderer.
Unchanged package member payloads remain byte-identical. Native application
behavior, font availability and routing are separate checks.
"""
import copy
import difflib
import hashlib
import io
import math
from pathlib import Path
import posixpath
from zipfile import ZipFile
from lxml import etree as E

NS = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
OBJECTS = {"sp", "grpSp", "cxnSp", "pic", "graphicFrame"}
PT = 12700
IDENTITY = (1., 0., 0., 1., 0., 0.)


def digest(value):
    return hashlib.sha256(value).hexdigest()


def xml(value):
    return E.fromstring(value, E.XMLParser(resolve_entities=False, no_network=True, remove_blank_text=False))


def canon(el):
    return E.tostring(el, method="c14n")


def kind(el):
    return E.QName(el).localname


def multiply(m, n):
    a, b, c, d, e, f = m
    g, h, i, j, k, l = n
    return (a*g+c*h, b*g+d*h, a*i+c*j, b*i+d*j, a*k+c*l+e, b*k+d*l+f)


def point(m, x, y):
    a, b, c, d, e, f = m
    return [a*x+c*y+e, b*x+d*y+f]


def transform(el, group=False):
    xf = el.find("p:grpSpPr/a:xfrm" if group else "p:spPr/a:xfrm", NS)
    if xf is None and kind(el) == "graphicFrame":
        xf = el.find("p:xfrm", NS)
    if xf is None:
        return None, None
    try:
        off, ext = xf.find("a:off", NS), xf.find("a:ext", NS)
        x, y, w, h = [int(v)/PT for v in (off.get("x"), off.get("y"), ext.get("cx"), ext.get("cy"))]
        angle = math.radians(int(xf.get("rot", "0"))/60000)
        flipx = -1 if xf.get("flipH") in ("1", "true") else 1
        flipy = -1 if xf.get("flipV") in ("1", "true") else 1
        center = (x+w/2, y+h/2)
        rotation = (math.cos(angle), math.sin(angle), -math.sin(angle), math.cos(angle), 0., 0.)
        around = multiply((1.,0.,0.,1.,*center), multiply(rotation, (flipx,0.,0.,flipy,-flipx*center[0],-flipy*center[1])))
        if group:
            co, ce = xf.find("a:chOff", NS), xf.find("a:chExt", NS)
            cx, cy, cw, ch = [int(v)/PT for v in (co.get("x"), co.get("y"), ce.get("cx"), ce.get("cy"))]
            if cw <= 0 or ch <= 0:
                return None, [x,y,w,h]
            child = (w/cw,0.,0.,h/ch,x-cx*w/cw,y-cy*h/ch)
            return multiply(around,child), [x,y,w,h]
        return around, [x,y,w,h]
    except (AttributeError, ValueError, TypeError, ZeroDivisionError):
        return None, None


def shallow(el):
    clone = copy.deepcopy(el)
    if kind(clone) == "grpSp":
        for child in list(clone):
            if kind(child) in OBJECTS:
                clone.remove(child)
    return canon(clone)


def records(root):
    result = []

    def walk(parent, owners=(), matrix=IDENTITY):
        for el in parent:
            if kind(el) not in OBJECTS:
                continue
            nv_path = {"sp":"nvSpPr", "grpSp":"nvGrpSpPr", "cxnSp":"nvCxnSpPr", "pic":"nvPicPr", "graphicFrame":"nvGraphicFramePr"}[kind(el)]
            nv = el.find("p:"+nv_path+"/p:cNvPr", NS)
            if nv is None:
                raise ValueError("Native object is missing its identity.")
            obj = {"id": nv.get("id"), "name": nv.get("name", ""), "kind": kind(el),
                   "group_path": list(owners), "sha256": digest(canon(el)), "element": el}
            group = kind(el) == "grpSp"
            local, bounds = transform(el, group)
            obj["local_bounds_pt"] = bounds
            if matrix is not None and bounds is not None:
                # Group frame bounds use the group's own rotation, not child-space scaling.
                frame = transform(el, False)[0] if not group else None
                if group:
                    clone = copy.deepcopy(el)
                    clone.tag = "{"+NS["p"]+"}sp"
                    prop = clone.find("p:grpSpPr", NS)
                    if prop is not None:
                        prop.tag = "{"+NS["p"]+"}spPr"
                    frame = transform(clone)[0]
                if frame is not None:
                    world = multiply(matrix,frame)
                    x,y,w,h = bounds
                    corners = [point(world,xx,yy) for xx,yy in ((x,y),(x+w,y),(x+w,y+h),(x,y+h))]
                    obj["world_bounds_pt"] = [min(c[0] for c in corners),min(c[1] for c in corners),
                                              max(c[0] for c in corners)-min(c[0] for c in corners),
                                              max(c[1] for c in corners)-min(c[1] for c in corners)]
                else:
                    obj["world_bounds_pt"] = None
            else:
                obj["world_bounds_pt"] = None
            obj["shallow_sha256"] = digest(shallow(el))
            result.append(obj)
            if group:
                walk(el, owners+(obj["id"],), multiply(matrix,local) if matrix is not None and local is not None else None)

    tree = root.find("p:cSld/p:spTree", NS)
    if tree is None:
        raise ValueError("Slide has no shape tree.")
    walk(tree)
    return result


class Package:
    def __init__(self, source_path):
        self.path = Path(source_path).resolve()
        if self.path.suffix.lower() != ".pptx":
            raise ValueError("Only standard .pptx files are supported by scoped patching.")
        self.data = self.path.read_bytes()
        self.sha256 = digest(self.data)
        with ZipFile(io.BytesIO(self.data)) as z:
            self.infos = z.infolist()
            names = z.namelist()
            if len(names) != len(set(names)):
                raise ValueError("Duplicate package member names.")
            if any(n.startswith("/") or ".." in n.split("/") or "\\" in n for n in names):
                raise ValueError("Unsafe package member name.")
            self.parts = {n:z.read(n) for n in names}
            self.comment = z.comment
        presentation = xml(self.parts["ppt/presentation.xml"])
        rels = {r.get("Id"):r for r in xml(self.parts["ppt/_rels/presentation.xml.rels"])}
        self.slides = []
        for slide in presentation.findall("p:sldIdLst/p:sldId", NS):
            rel = rels[slide.get("{"+NS["r"]+"}id")]
            if rel.get("TargetMode") == "External":
                raise ValueError("Slide relationship cannot be external.")
            self.slides.append(posixpath.normpath(posixpath.join("ppt",rel.get("Target"))).lstrip("/"))

    def slide(self, index):
        if not isinstance(index,int) or isinstance(index,bool) or not 1 <= index <= len(self.slides):
            raise ValueError("slide_index is outside presentation order.")
        part = self.slides[index-1]
        return part, xml(self.parts[part])


def audit_connections(source_path, slide_indices=None, include_objects=False):
    package = Package(source_path)
    slides = slide_indices if slide_indices is not None else list(range(1,len(package.slides)+1))
    if not isinstance(slides,list) or not slides:
        raise ValueError("slide_indices must be a nonempty list.")
    result = {"source_sha256":package.sha256,"passed":True,"slides":[],
              "limits":"Recursive native reference/site audit, not route collision, rendered clarity or native application certification."}
    for index in slides:
        _, root = package.slide(index)
        items = records(root)
        byid = {i["id"]:i for i in items}
        errors, warnings, edges = [], [], []
        if len(byid) != len(items):
            errors.append({"category":"duplicate_id"})
        names = {}
        for obj in items:
            names.setdefault(obj["name"],0); names[obj["name"]] += 1
            if not obj["id"] or not obj["id"].isascii() or not obj["id"].isdigit() or len(obj["id"]) > 10 or not 0 < int(obj["id"]) <= 4294967295:
                errors.append({"category":"invalid_object_id","shape_name":obj["name"]})
            if obj["world_bounds_pt"] is None:
                warnings.append({"shape_id":obj["id"],"category":"world_geometry_unresolved"})
        for name,count in names.items():
            if count>1:
                warnings.append({"shape_name":name,"category":"ambiguous_name_use_id"})
        for obj in items:
            if obj["kind"] != "cxnSp":
                continue
            edge = {"id":obj["id"],"name":obj["name"],"group_path":obj["group_path"],"valid":True}
            for tag,key in (("stCxn","source"),("endCxn","target")):
                binding = obj["element"].find("p:nvCxnSpPr/p:cNvCxnSpPr/a:"+tag,NS)
                target = byid.get(binding.get("id")) if binding is not None else None
                site = binding.get("idx") if binding is not None else None
                if target is None or site is None or not site.isascii() or not site.isdigit() or len(site) > 10 or target["kind"] == "cxnSp":
                    edge["valid"] = False
                    errors.append({"shape_id":obj["id"],"end":key,"category":"invalid_endpoint_reference"})
                    continue
                edge[key] = {"id":target["id"],"site_index":int(site),"group_path":target["group_path"]}
                custom = target["element"].find("p:spPr/a:custGeom/a:cxnLst",NS)
                preset = target["element"].find("p:spPr/a:prstGeom",NS)
                count = len(custom) if custom is not None else 4 if preset is not None and preset.get("prst") in ("rect","ellipse") else None
                if count is not None and int(site) >= count:
                    edge["valid"] = False
                    errors.append({"shape_id":obj["id"],"end":key,"category":"invalid_connection_site","site_count":count})
                elif count is None:
                    warnings.append({"shape_id":obj["id"],"end":key,"category":"connection_site_count_unresolved"})
            edge["cross_group"] = "source" in edge and "target" in edge and edge["source"]["group_path"] != edge["target"]["group_path"]
            edge["relationship_scope"] = "cross_group" if edge["cross_group"] else "group_internal" if edge.get("source",{}).get("group_path") else "top_level"
            edges.append(edge)
        row = {"slide_index":index,"object_count":len(items),"group_count":sum(i["kind"]=="grpSp" for i in items),
               "shape_count":sum(i["kind"]=="sp" for i in items),"connector_count":len(edges),
               "valid_connector_count":sum(e["valid"] for e in edges),"cross_group_connector_count":sum(e["cross_group"] for e in edges),
               "group_internal_connector_count":sum(e["relationship_scope"]=="group_internal" for e in edges),
               "errors":errors,"warnings":warnings,"connectors":edges}
        if include_objects:
            row["objects"] = [{k:v for k,v in obj.items() if k != "element"} for obj in items]
        result["slides"].append(row)
        result["passed"] = result["passed"] and not errors
    return result


def segments(shape):
    body = shape.find("p:txBody",NS)
    if body is None or body.find(".//a:fld",NS) is not None:
        raise ValueError("Text patch requires an existing text frame without dynamic fields.")
    groups = []
    for paragraph in body.findall("a:p",NS):
        current = []
        for child in paragraph:
            if child.tag == "{"+NS["a"]+"}br":
                groups.append(current); current=[]
            elif child.tag == "{"+NS["a"]+"}r":
                current.extend(child.findall("a:t",NS))
        groups.append(current)
    return groups


def replace_segment(nodes, text):
    old = "".join(n.text or "" for n in nodes)
    if old == text:
        return
    if not nodes:
        raise ValueError("Cannot add text to an empty segment without an explicit style design.")
    owners = []
    for i,n in enumerate(nodes):
        owners.extend([i]*len(n.text or ""))
    values = ["" for _ in nodes]
    for op,i,j,k,l in difflib.SequenceMatcher(None,old,text,autojunk=False).get_opcodes():
        if op == "equal":
            for pos,char in zip(range(i,j),text[k:l]):
                values[owners[pos]] += char
        elif op in ("replace","insert"):
            owner = owners[i] if i<len(owners) else owners[-1] if owners else 0
            values[owner] += text[k:l]
    for node,value in zip(nodes,values):
        node.text = value


def masked(el, edit):
    clone = copy.deepcopy(el)
    if "text" in edit:
        for node in clone.findall("p:txBody//a:t",NS):
            node.text = ""
    xf = clone.find("p:grpSpPr/a:xfrm" if kind(clone)=="grpSp" else "p:spPr/a:xfrm",NS)
    for field,tag,attr in (("left","off","x"),("top","off","y"),("width","ext","cx"),("height","ext","cy")):
        if field in edit and xf is not None:
            node = xf.find("a:"+tag,NS)
            if node is not None:
                node.set(attr,"0")
    return canon(clone)


def prepare_patch(package, edits):
    if not isinstance(edits,list) or not 1 <= len(edits) <= 1000:
        raise ValueError("Provide 1–1000 explicit scoped edits.")
    if any(n.startswith("_xmlsignatures/") or n.lower().endswith("vbaproject.bin") for n in package.parts):
        raise ValueError("Signed or macro-bearing packages require a separate preservation workflow.")
    roots, touched, plan = {}, set(), []
    for edit in edits:
        if set(edit)-{"slide_index","shape_name","shape_id","expected_shape_sha256","text","left","top","width","height","coordinate_space"}:
            raise ValueError("Unsupported edit field; scoped patch cannot silently ignore settings.")
        index = edit.get("slide_index")
        if index not in roots:
            part,root = package.slide(index); roots[index]=(part,root)
        part,root = roots[index]
        items = records(root)
        hasid,hasname = "shape_id" in edit,"shape_name" in edit
        if hasid == hasname:
            raise ValueError("Select exactly one shape_id or shape_name.")
        matches = [v for v in items if v["id"]==str(edit["shape_id"])] if hasid else [v for v in items if v["name"]==edit["shape_name"]]
        if len(matches)!=1:
            raise ValueError("Shape selector is missing or ambiguous; use a unique shape_id.")
        obj = matches[0]; el = obj["element"]
        key=(index,obj["id"])
        if key in touched:
            raise ValueError("Combine changes to one object in one edit.")
        if any(row["slide_index"] == index and (row["shape_id"] in [int(v) for v in obj["group_path"]] or int(obj["id"]) in [int(v) for v in row["group_path"]]) for row in plan):
            raise ValueError("Ancestor and descendant edits must be planned in separate patches.")
        touched.add(key)
        if edit.get("expected_shape_sha256") and edit["expected_shape_sha256"]!=obj["sha256"]:
            raise ValueError("Object conflict: selected shape changed since planning.")
        fields = set(edit)&{"text","left","top","width","height"}
        if not fields:
            raise ValueError("Edit has no supported mutation.")
        original = copy.deepcopy(el)
        if "text" in fields:
            value = edit["text"]
            if not isinstance(value,str) or len(value)>50000 or "\r" in value:
                raise ValueError("text must be a bounded string using LF separators.")
            groups=segments(el); lines=value.split("\n")
            if len(groups)!=len(lines):
                raise ValueError("Changing paragraph/soft-break count requires an explicit text-layout operation.")
            for nodes,line in zip(groups,lines):
                replace_segment(nodes,line)
        if fields-{"text"}:
            if obj["group_path"] and edit.get("coordinate_space")!="local":
                raise ValueError("Grouped geometry requires coordinate_space=local; values are parent-space points.")
            xf=el.find("p:grpSpPr/a:xfrm" if obj["kind"]=="grpSp" else "p:spPr/a:xfrm",NS)
            if xf is None:
                raise ValueError("Selected geometry type is unsupported.")
            for field,tag,attr in (("left","off","x"),("top","off","y"),("width","ext","cx"),("height","ext","cy")):
                if field in edit:
                    value=edit[field]
                    if isinstance(value,bool) or not isinstance(value,(int,float)) or not math.isfinite(value) or abs(value)>100000 or field in ("width","height") and value<=0:
                        raise ValueError("Invalid geometry in parent-space points.")
                    node=xf.find("a:"+tag,NS)
                    if node is None:
                        raise ValueError("Geometry transform is incomplete.")
                    node.set(attr,str(round(value*PT)))
        if masked(original,edit)!=masked(el,edit):
            raise ValueError("Patch exceeded its authorized field scope.")
        plan.append({"slide_index":index,"shape_id":int(obj["id"]),"shape_name":obj["name"],"fields":sorted(fields),
                     "before_sha256":obj["sha256"],"after_sha256":digest(canon(el)),"group_path":obj["group_path"]})
    changed={}
    untouched=0
    for index,(part,root) in roots.items():
        original=xml(package.parts[part]); old={v["id"]:v for v in records(original)}; new={v["id"]:v for v in records(root)}
        if len(old)!=len(records(original)) or set(old)!=set(new):
            raise ValueError("Duplicate identities or object inventory changed.")
        for objid,oldobj in old.items():
            if (index,objid) not in touched:
                if oldobj["shallow_sha256"]!=new[objid]["shallow_sha256"]:
                    raise ValueError("Unselected object changed.")
                untouched+=1
        # All outside-object slide data (background, extensions, timing, etc.)
        # must also survive; targeted objects are removed only for this check.
        def skeleton(value):
            clone=copy.deepcopy(value); tree=clone.find("p:cSld/p:spTree",NS)
            for obj in list(tree):
                if kind(obj) in OBJECTS:tree.remove(obj)
            return canon(clone)
        if skeleton(original)!=skeleton(root):
            raise ValueError("Unselected slide structure changed.")
        if canon(root)!=canon(original):
            changed[part]=E.tostring(root,encoding="UTF-8",xml_declaration=True,standalone=True)
    return changed,plan,untouched


def plan_local_patch(source_path, edits):
    package=Package(source_path)
    changed,plan,untouched=prepare_patch(package,edits)
    normalized=[]
    for edit,row in zip(edits,plan):
        value=dict(edit);value.pop("shape_name",None);value["shape_id"]=row["shape_id"]
        value["expected_shape_sha256"]=row["before_sha256"];normalized.append(value)
    return {"source_sha256":package.sha256,"edits":normalized,"plan":plan,"changed_parts":sorted(changed),
            "untouched_object_count":untouched,"write_performed":False,
            "limits":"Planned local patch only. Preserve paragraph/break count and run properties. Geometry uses parent-space points; routing and fit require native review."}


def apply_local_patch(source_path, output_path, expected_source_sha256, edits):
    package=Package(source_path)
    if package.sha256!=expected_source_sha256:
        raise ValueError("Source conflict: file changed since planning. Reinspect and plan against the current file.")
    output=Path(output_path).resolve()
    if output.suffix.lower() != ".pptx":
        raise ValueError("Output must use the .pptx extension.")
    if output==package.path or output.exists():
        raise ValueError("Scoped patch requires a new output path; source and existing files are never overwritten.")
    changed,plan,untouched=prepare_patch(package,edits)
    buffer=io.BytesIO()
    with ZipFile(buffer,"w") as out:
        out.comment=package.comment
        for info in package.infos:
            out.writestr(info,changed.get(info.filename,package.parts[info.filename]))
    data=buffer.getvalue()
    with ZipFile(io.BytesIO(data)) as candidate:
        if set(candidate.namelist())!=set(package.parts):
            raise ValueError("Package member inventory changed.")
        preserved=sum(1 for n,b in package.parts.items() if n not in changed and candidate.read(n)==b)
        if preserved!=len(package.parts)-len(changed):
            raise ValueError("Unselected package member changed.")
    if digest(package.path.read_bytes())!=package.sha256:
        raise ValueError("Source conflict: file changed during patch preparation.")
    output.parent.mkdir(parents=True,exist_ok=True)
    with output.open("xb") as f:
        f.write(data)
    return {"source_sha256":package.sha256,"output_sha256":digest(data),"output_path":str(output),"plan":plan,
            "changed_parts":sorted(changed),"unchanged_part_count":preserved,"untouched_object_count":untouched,
            "preservation_passed":True,"native_review_required":True,
            "limits":"Unchanged package payloads and unselected object XML preserved. No file lock, signature renewal, automatic rerouting or native application certification."}
