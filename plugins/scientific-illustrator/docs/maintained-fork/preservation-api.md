# Recursive connections and scoped local updates

These are **file-only OOXML tools**. They operate on an explicitly named local PPTX, independently of the active presentation and session metadata. No application is opened and no network call is made by these operations. Keep private inputs, plans, results and screenshots outside the source repository.

## Actual connector references

`powerpoint_audit_connections` takes `source_path`, optional `slide_indices` in presentation order, and optional `include_objects`. It recursively reads native objects inside groups and nested groups. Every connector is checked against actual `a:stCxn` / `a:endCxn` IDs and zero-based connection-site indices. A link between different endpoint group paths is counted as cross-group; group membership never makes an otherwise valid reference missing.

Duplicate IDs, missing/non-shape endpoints and known-invalid site indices fail the gate. Duplicate names are a warning to select by native ID. The audit checks four sites for rectangles/ellipses and explicit custom-geometry connection lists. Other presets or group connection sites remain **unresolved coverage warnings**, rather than guessed successes or false failures. Local bounds and composed world bounding boxes include nested scale, rotation and flips. Missing transforms produce an explicit warning. A valid reference does not establish the chosen site's semantic direction, automatic rerouting, collision-free paths or visual clarity.

The OOXML `powerpoint_audit_figure` includes this recursive reference report. Its older top-level route/text heuristics remain separate and incomplete for grouped geometry. COM and Office.js live audits have not acquired this new recursive implementation.

## Plan, then apply

1. Inspect the current actual input and back it up locally.
2. Call `powerpoint_plan_local_patch` with `source_path` and explicit `edits`. Planning is read-only. Each edit selects exactly one `shape_id` or unique `shape_name` plus `slide_index`.
3. Review the returned affected objects, fields and package parts. Retain the returned `source_sha256` and normalized `edits` with object preconditions.
4. Call `powerpoint_apply_local_patch` with the same source, a **new** `output_path`, `expected_source_sha256` and those normalized edits. Stale sources/objects and existing outputs are rejected; re-inspect user changes and make a new plan. Do not silently reset preconditions or overwrite the current deck.
5. Audit the candidate, inspect its rendering and test native editing/open/reopen separately. After any native save, recompute hashes and re-audit against the new bytes. Apply authorized delivery only after those separate gates.

A synthetic example edit:

```json
{
  "slide_index": 1,
  "shape_name": "Method node",
  "text": "Revised method\nRevised result",
  "width": 110,
  "coordinate_space": "local"
}
```

Geometry values are **points in the immediate parent coordinate space**, including for nested groups. Grouped geometry changes require explicit `coordinate_space: "local"`; world coordinates are not implicitly inverted. Group frame movement does not rewrite child geometry. Editing an ancestor and descendant in the same batch is rejected: split into separately inspected plans.

## Preservation contract and limits

- Only explicitly selected text contents or `left`, `top`, `width`, `height` attributes change. Object IDs, grouping, ordering, existing text run/paragraph properties and connector binding XML survive.
- Text changes retain the existing paragraph/soft-break count and run structure. Inserted/replaced characters inherit the neighboring run's properties. Adding empty text segments, dynamic fields or new paragraph structure requires a separate explicit layout operation and is rejected here.
- Unselected objects' canonical XML is checked; for ancestor groups their own XML is checked excluding child objects. Outside-object slide data is checked separately. Only changed slide parts are serialized. Every other **uncompressed package member payload** is checked byte-for-byte, including themes, layouts, notes, media and unknown parts. ZIP container bytes may differ; there is no claim of identical compression streams.
- Signed or macro-bearing inputs are rejected. The tool does not renew signatures, supply a file lock, migrate fonts, repair arbitrary packages, resize text automatically or reroute connectors. A final source hash check catches changes during preparation, but is not transactional locking against a simultaneous writer at publication.
- `native_review_required` stays true. OOXML bindings and saved XML do not prove behavior in desktop PowerPoint. Native saves may reserialize package parts, and do not inherit the scoped patch's untouched-payload claim.
- Legacy `powerpoint_update_shape`, `powerpoint_save` and renderer-based workflows retain their previous behavior. Use the new plan/apply API when user-edit/package preservation is required; do not attach its guarantees to every tool.

Run `node --run test:package-preservation` for synthetic nested/cross-group cases, transformed coordinates, conflict rejection, all-or-nothing failure, run-property preservation and opaque-part preservation. This test also calls the public MCP endpoints without launching a desktop application.
