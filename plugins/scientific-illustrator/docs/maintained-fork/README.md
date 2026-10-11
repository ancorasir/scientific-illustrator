# Maintained fork: editable research figures

This fork builds on [Scientific Illustrator](https://github.com/icebird1998/scientific-illustrator). Retain the upstream MIT license and attribution. The first extension targets editable, information-dense diagrams for documents and presentations. It is a development extension, not a new stable release or certification of every platform.

## What is implemented

- A read-only `powerpoint_resolve_figure_style` MCP tool resolves optional typography roles, native line widths and spacing tokens at the **final cropped figure size**.
- Presets: `print-standard`, `print-dense`, `presentation`. MiSans is a selectable preset family, not an implicit migration of user documents; `font_family` overrides it.
- Source-to-output scale compensation uses cropped figure width, not slide width. The height budget requires explicit figure, body, caption and paragraph-gap dimensions; missing data is reported as unchecked.
- `line_spacing_pt`, `paragraph_space_before_pt`, `paragraph_space_after_pt` are implemented in the OOXML working-copy backend. Explicit font assignment also sets Latin, East Asian and complex-script run fields.
- Other backends reject explicit paragraph settings **before dispatch**. This includes settings nested in sequence/table operations; unsupported parameters must not appear to succeed.
- Synthetic tests cover style scaling, invalid inputs, budget overflow, backend guards, the public MCP endpoint, saved OOXML typography and preservation of an unselected object.

These changes do not supply a constraint solver, automatic text measurement, collision-free routing or native application verification. The existing installed upstream plugin is not changed by cloning this fork.

## Use in a drawing session

1. Follow the [layout guide](layout-guide.md). Inspect and back up the current actual source; work in a copy.
2. Call `powerpoint_status` and `powerpoint_get_capabilities`. For exact paragraph settings choose `powerpoint_set_backend(ooxml)` before any mutation. If the session is already locked to another backend, use a fresh session.
3. Call `powerpoint_resolve_figure_style` for each text role. Resolve styles once per role and use the same returned values for repeated nodes.
4. Spread `text_style` into `powerpoint_add_shape`, `powerpoint_add_textbox` or an explicitly selected `powerpoint_update_shape`. Spread `line_style` into native frames/connectors. Do not pass the result wrapper as a drawing command.
5. Construct bound connectors where supported, reserve routes and inspect at final size. Keep document captions outside the drawing when appropriate.
6. Run structure/content/style checks, visual review and separate native editing tests. A pure budget or XML test is not a native-open or visual-fit result.

Example resolver input, using generic dimensions:

```json
{
  "profile": "print-dense",
  "role": "body",
  "final_width_mm": 140,
  "source_figure_width_mm": 140,
  "final_height_mm": 210,
  "body_width_mm": 160,
  "body_height_mm": 240,
  "caption_height_mm": 8,
  "paragraph_gap_mm": 5
}
```

This returns a 10 pt body font with 12.5 pt exact leading, fixed padding and no shrink-to-fit. A source figure twice as wide receives twice these source sizes so that insertion restores the requested final sizes. Font availability and real text fit still need rendering inspection. Width and height warnings do not certify the document's actual pagination.

## Development and verification

Run `node --run test:figure-style` for the new regression suite and `node --run test` for the full existing suite. Python requires `python-pptx`; set `SCIENTIFIC_ILLUSTRATOR_PYTHON` if the Python executable is not on PATH. The typography test saves/reopens a synthetic deck in python-pptx; it does **not** open desktop PowerPoint.

The [upgrade roadmap](roadmap.md) separates shipped functionality from planned work. Keep source changes on reviewed topic branches, track upstream independently and retain rollback points. The fork-specific extension is not published under a separate installed-plugin identity yet; do not replace a plugin cache or claim the desktop MCP tools have switched automatically. Explicit local-source runs can use this fork's server now. A separately named distribution and migration check belong to the release phase.

## Privacy boundary

Only generalized code, rules and synthetic fixtures belong in this repository. Keep real proposals, unpublished research, personal details, local paths, screenshots, recordings, native receipts, hashes of private files and source inventories outside it. Even a repository branch is public when the repository is public.

Rewrite general lessons rather than copying private preparation files. Stage an explicit file list, inspect every added line and all metadata, then run `node scripts/check-public-diff.mjs`. That guard detects common mistakes; it cannot prove material is public or replace human review. Use an account's public/noreply identity in commit metadata. Do not open public issues or send upstream private fixtures automatically.
