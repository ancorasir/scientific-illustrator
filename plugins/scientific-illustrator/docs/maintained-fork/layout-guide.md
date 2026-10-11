# Research figure layout guide

These are adjustable working rules, not mandatory journal standards. A successful diagram preserves meaning, native editability and the user's existing edits while remaining readable at its actual destination size.

## Size before layout

Measure the destination document's page, margins, available text area, caption and paragraph gaps. Do not assume a universal image width or use an oversized slide as evidence of readability. Define final cropped bounds separately from slide bounds. Fonts, strokes, arrowheads and spacing all scale on insertion. Titles can be stored in notes when the document supplies its own caption.

The profile resolver takes explicit dimensions; it does not read a Word document. A page proof checks space and scale; actual pagination also depends on surrounding paragraphs and must be tested separately when editing the document.

## Typography hierarchy

Prefer one approved font family and regular/bold weights. Define section, task, body and note roles explicitly. Do not change a user's existing template or font inheritance to satisfy a preset.

|Role|Standard print preset|Dense print preset|Leading|
|---|---:|---:|---:|
|Section band|12 pt|11.5 pt|1.2×|
|Task or key node|12 pt|10.5 pt|1.2×|
|Methods and explanation|12 pt|10 pt|1.25×|
|Supporting note|10.5 pt|9.5 pt|1.25×|

Treat these values as starting points. A planned, role-specific size reduction is different from automatic shrink-to-fit. Reflow and enlarge relevant node classes before reducing secondary text. Do not use the space gained by smaller type to pack in more content. Font-aware measurement is useful but does not prove native glyph fit or line wrapping.

## Space and geometry

- Multiline method nodes usually need horizontal padding around 1.5–2 mm and vertical padding around 1–1.8 mm at final size. Bands and unframed labels can use less; one padding value does not suit every object.
- Reserve roughly 2–3 mm for short connector lanes and 3–6 mm between print modules, adjusted after final-size review. Keep line separation visible even when background colors match.
- Use repeated node classes and aligned columns/baselines. Module heights may differ when their content differs; forced equal heights can waste space or compress text.
- Use a consistent stroke width (the print preset uses 0.6 pt) and restrained arrowheads. Distinguish main flow and feedback with route, direction, dash and semantic color; avoid arbitrary changes in thickness.
- Keep contrast high. Use a small palette consistently: dark text, pale panels and optional colored section bands. Color supports labels and structure; it does not replace them.

## Connections and editable structure

Build semantic nodes, named relationships and ports before positioning shapes. Bind relationship endpoints to real objects in a supported backend. Free arrow segments are appropriate for axes or decoration, not substitutes for bound relationships. Keep cross-column and feedback corridors clear.

Group only logical regions that should move together. Keep child text and shapes editable. Endpoint binding does not supply global alignment constraints, edge-label attachment or automatic collision avoidance. Recheck routes and labels after movement, resizing or grouping. Office.js's current geometry-backed routing must not be described as native endpoint binding.

## Independent acceptance gates

1. **Content and preservation:** compare meaning and wording against the latest actual source. Planned line breaks differ from omissions or unauthorized rewriting. Preserve untouched pages, objects, notes, masters and package dependencies.
2. **Structure and typography:** check IDs, connector targets/sites, explicit font fields, stroke widths, leading, padding, grouping and source/output scaling. Traverse grouped children; top-level-only checks miss real endpoints.
3. **Visual:** inspect hierarchy, whitespace, focus, text fit, routes and captions at final physical size. No overflow does not establish good layout.
4. **Native behavior:** in a disposable copy, move a multi-edge node/group, resize a different node and increase text. Save and reopen in the real application. Check endpoint following, unchanged typography and untouched objects. Record the exact final file version separately from automated tests.

If a gate is unavailable or uncertain, report that limitation instead of replacing it with another kind of check. Never flatten a whole diagram to hide an editing defect.

## External reading

- [Nature Reviews artwork guide](https://www.nature.com/documents/natrev-artworkguide.pdf)
- [Nature Methods Points of View](https://blogs.nature.com/blog/tag/points-of-view/)
- [PLOS figure guidelines](https://journals.plos.org/plosone/s/figures)

These references inform general clarity and publication preparation. The exact role sizes and spacing above are configurable presets, not attributed universal requirements.
