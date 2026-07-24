# @stll/premirror-bridge

A **separately-versioned folio plugin**: a
[`@chenglou/pretext`](https://www.npmjs.com/package/@chenglou/pretext)-backed
`SegmentFitEngine` for folio-core's measurement seam, with a frozen parity
suite against folio's legacy word-walk.

It depends only on `@stll/folio-core` (the seam, shipped in
`layout-engine/measure/segmentFit.ts`) plus its own third-party dep,
`@chenglou/pretext` — the only package in this workspace allowed to import it
(the one-pretext-surface invariant). It is vendored as a folio workspace member
(`packages/premirror-bridge`), resolved via `workspace:*` instead of the former
machine-local `bun link`.

## Usage

```ts
import { pretextSegmentFitEngine } from "@stll/premirror-bridge";
import { setSegmentFitEngine } from "@stll/folio-core/layout-engine/measure/segmentFit";

setSegmentFitEngine(pretextSegmentFitEngine);
globalThis.__folioFeatureFlags = { segmentFitLineBreaking: true };
```

With the flag off or the engine not installed, folio measures exactly as before
(the seam is dormant). Turning it on routes plain-text line fitting through
pretext's prepare-once/fit-by-arithmetic model.

## Dependency pin (workspace vs. release)

Inside this monorepo the package resolves `@stll/folio-core` via
`workspace:*`. If it is ever published standalone, replace the workspace range
with an exact version pin of the folio-core release that carries the
segment-fit seam. That is the only ship-time change; the source does not
move.

## Scope

This package is E-2: the pretext engine + parity harness. The composer golden
fixtures that pull `@premirror/*` belong to E-3 (the `@premirror/*` rebase) and
are intentionally not included here — keeping E-2's dependency surface to
`@chenglou/pretext` + `@stll/folio-core` only.

## Credit

See [NOTICE](./NOTICE). The engine wraps `@chenglou/pretext` (MIT); the bridge
belongs to the premirror line, `samwillis/premirror` (MIT © 2026 Sam Willis).
