import assert from "node:assert/strict";
import { execFile, spawn } from "node:child_process";
import { promisify } from "node:util";
import { createInterface } from "node:readline";
import { fileURLToPath } from "node:url";
import { mkdtempSync, writeFileSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import { execFileSync } from "node:child_process";
import { resolveFigureStyle, assertParagraphSupport } from "../plugins/scientific-illustrator/scripts/figure-style.mjs";

const dense = resolveFigureStyle({profile:"print-dense",role:"body",final_width_mm:140});
assert.equal(dense.text_style.font_size,10);
assert.equal(dense.text_style.line_spacing_pt,12.5);
assert.equal(dense.text_style.text_autofit,"none");
const scaled = resolveFigureStyle({profile:"print-dense",role:"body",final_width_mm:140,source_figure_width_mm:280});
assert.equal(scaled.text_style.font_size,20);
assert.equal(scaled.text_style.line_spacing_pt,25);
assert.equal(scaled.line_style.line_width,1.2);
assert.equal(scaled.spacing_mm.moduleGap,6);
assert.equal(scaled.text_style.margin_top * scaled.insertion_scale,dense.text_style.margin_top);
assert.throws(()=>resolveFigureStyle({final_width_mm:140,source_figure_width_mm:1}),/API limits/);
for (const args of [{final_width_mm:0},{final_width_mm:NaN},{final_width_mm:140,profile:"unknown"},{final_width_mm:140,role:"unknown"},{final_width_mm:140,leading_ratio:.9},{final_width_mm:140,font_family:""}]) {
  assert.throws(()=>resolveFigureStyle(args));
}
const overflow = resolveFigureStyle({final_width_mm:140,final_height_mm:230,body_width_mm:130,body_height_mm:240,caption_height_mm:8,paragraph_gap_mm:5});
assert.equal(overflow.height_budget.fits,false);
assert.equal(overflow.height_budget.remaining_mm,-3);
assert.equal(overflow.warnings.length,2);
assert.equal(resolveFigureStyle({final_width_mm:140,final_height_mm:230}).height_budget.checked,false);
assertParagraphSupport("ooxml",{line_spacing_pt:12.5});
assertParagraphSupport("officejs",{font_size:12});
assert.throws(()=>assertParagraphSupport("ooxml",{line_spacing_pt:0}),/finite/);
assert.throws(()=>assertParagraphSupport("ooxml",{paragraph_space_after_pt:-1}),/finite/);
for (const backend of ["officejs","com"]) {
  assert.throws(()=>assertParagraphSupport(backend,[{cells:[{line_spacing_pt:12.5}]}]),/OOXML/);
}

// Exercise the public MCP endpoint without opening or modifying any application.
const server = fileURLToPath(new URL("../plugins/scientific-illustrator/scripts/powerpoint-server.mjs",import.meta.url));
const child = spawn(process.execPath,[server],{stdio:["pipe","pipe","pipe"]});
const lines = createInterface({input:child.stdout});
const responses = new Map();
let stderr=""; child.stderr.on("data",s=>stderr+=s);
try {
  const completed = new Promise((resolve,reject)=>{
    const timer=setTimeout(()=>reject(new Error(`MCP timeout: ${stderr}`)),8000);
    lines.on("line",line=>{
      try {
        const r=JSON.parse(line);responses.set(r.id,r);
        if(responses.size===2){clearTimeout(timer);resolve();}
      } catch(e){clearTimeout(timer);reject(e);}
    });
    child.once("error",reject);
  });
  child.stdin.write(JSON.stringify({jsonrpc:"2.0",id:1,method:"tools/list",params:{}})+"\n");
  child.stdin.write(JSON.stringify({jsonrpc:"2.0",id:2,method:"tools/call",params:{name:"powerpoint_resolve_figure_style",arguments:{profile:"print-dense",role:"body",final_width_mm:140}}})+"\n");
  await completed;
  const tools=responses.get(1).result.tools;
  assert(tools.some(t=>t.name==="powerpoint_resolve_figure_style"));
  for(const name of ["powerpoint_add_textbox","powerpoint_add_shape","powerpoint_update_shape"]){
    assert(tools.find(t=>t.name===name).inputSchema.properties.line_spacing_pt);
  }
  assert.equal(responses.get(2).result.structuredContent.text_style.line_spacing_pt,12.5);
} finally { lines.close();child.kill(); }

const python = process.env.SCIENTIFIC_ILLUSTRATOR_PYTHON || (process.platform==="win32"?"python":"python3");
const {stdout} = await promisify(execFile)(python,[fileURLToPath(new URL("./figure-typography-test.py",import.meta.url))],{maxBuffer:1024*1024});
process.stdout.write(stdout);
// The publication guard uses generic staged fixtures in an isolated repository.
const guard=fileURLToPath(new URL("./check-public-diff.mjs",import.meta.url));
const fixture=mkdtempSync(path.join(os.tmpdir(),"public-diff-test-"));
try {
  const git = args => execFileSync("git",args,{cwd:fixture,stdio:"pipe"});
  const runGuard = () => execFileSync(process.execPath,[guard],{cwd:fixture,stdio:"pipe"});
  git(["init","--quiet"]);
  writeFileSync(path.join(fixture,"README.md"),"Generic public guidance\n");git(["add","README.md"]);runGuard();
  writeFileSync(path.join(fixture,"README.md"),"person"+"@"+"example.invalid\n");git(["add","README.md"]);assert.throws(runGuard);
  writeFileSync(path.join(fixture,"README.md"),"/"+"home/"+"example/"+"private.txt\n");git(["add","README.md"]);assert.throws(runGuard);
  writeFileSync(path.join(fixture,"README.md"),"Generic public guidance\n");git(["add","README.md"]);
  writeFileSync(path.join(fixture,"document.docx"),"not a real document");git(["add","document.docx"]);assert.throws(runGuard);
} finally { rmSync(fixture,{recursive:true,force:true}); }
console.log("Figure style presets, crop scaling, budget warnings, backend guards and MCP endpoint passed.");
