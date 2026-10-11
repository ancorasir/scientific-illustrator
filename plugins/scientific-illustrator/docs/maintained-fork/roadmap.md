# Upgrade roadmap and maintenance boundaries

## First increment: implemented

Keep the upstream drawing backends. Add optional role-based styles, cropped-output scale compensation, explicit height budgeting, OOXML paragraph leading/spacing and CJK font assignment. Wire these capabilities into the drawing skill, MCP API and synthetic regression suite. Do not label this development increment a stable release or claim that every backend supports it.

## Highest-value next increments

|Priority|Capability|Why it matters|Acceptance evidence|
|---|---|---|---|
|1|Recursive grouped-object and connector audit|A top-level scan misses child endpoints and can falsely report missing bindings|Synthetic nested groups, cross-group links, site-index failures; local vs world coordinates distinguished; rotated/flipped transforms explicit|
|2|Content-preserving local updates|Regeneration must not overwrite later user edits or original themes|Before/after inventories; changed-object scope; untouched package dependencies preserved; an explicit conflict report|
|3|Semantic FigureSpec and constrained layout|Stable node/edge IDs, named ports, repeated classes, columns and reserved lanes avoid ad hoc coordinates|Text growth and node moves trigger an explicit reflow/check; user-locked positions retained; no promise of arbitrary collision-free dragging|
|4|Font-aware capacity and final-size proof|Character-count heuristics miss CJK width, leading and padding|Measured glyph width plus declared wrap uncertainty; final crop/page proof; a separate native-render result|
|5|Native regression harness and release identity|File validity alone does not prove application editing; a fork must not silently masquerade as the installed upstream plugin|Platform capability matrix; isolated plugin/state identifiers; exact source/tag/version; install, upgrade and rollback checks|

Apply priorities to the current need rather than expanding all platforms at once. The first target is a local OOXML workflow. COM and Office.js paragraph interfaces require implementation and real-platform validation before changing the present rejection policy.

## Architecture direction

```text
Current source and authorized edit scope
        ↓
FigureSpec: semantic nodes, ports, relationships, user locks
        + FigureProfile: final dimensions, roles, spacing
        ↓
Layout constraints and font-aware capacity diagnostics
        ↓
Existing COM / OOXML / Office.js / draw.io adapters
        ↓
Content preservation + structure + visual + native checks
        ↓
Editable working copy and final-size preview
```

Storybook can later showcase a style editor or reusable UI components. It is not the solver for PPT endpoint binding, graph constraints or Office package preservation.

## Sustainable maintenance

- Keep upstream and fork remotes distinct; record the actual upstream commit for each integration. Retain MIT notices and original author attribution.
- Use small topic branches and reviewable diffs. Prefer generic public fixes that can be upstreamed separately when authorized. Do not edit an installed cache as the source of truth.
- Do not publish a release version until its source, tests, platform limits and installation behavior are recorded. Keep the prior working version available for rollback.
- Before separate distribution, assign a unique plugin/marketplace/state identity and adjust installation documentation and validation together. Do not change provenance fields or claim authorship of upstream work merely to rename a package.
- Keep regression data synthetic and repositories independent from private document/data stores. Do not add private materials even to a feature branch or public issue.
- Distinguish source unit tests, saved OOXML checks, renderer checks and observed native application behavior. Each test result establishes only its own scope.
