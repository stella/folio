# Folio seam architecture

Target architecture for folio: a layered set of packages joined by explicit,
typed seams, so that the same engine powers multiple framework adapters
(React, Vue), multiple hosts (web, server), and headless document editing.

This doc defines the technical seams. The higher-level ownership and
interoperability contract is defined in
[DOCX platform boundary](./docx-platform.md).

## Principle

Cut the seams by **reason to change** and **portability profile**, not by
"React or not". Today folio has one coarse seam — `core/` (React-free) vs the
React adapter — which is why the adapter reaches into ~91 core modules: the
boundary is in the wrong place. Replace it with layers where each boundary is a
typed contract.

## Responsibility map

From most-portable (bottom) to most-framework-specific (top).

| Layer           | Owns                                                                                                                                                                           | Profile                                |
| --------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ | -------------------------------------- |
| `model`         | Pure data shapes: `Document`, `Layout`, `FlowBlock`, `Measure`, content types, ids                                                                                             | Serializable data                      |
| `engine`        | Pure compute: OOXML read/write, `Document→FlowBlocks`, pagination/line-break, markdown, pure transforms (fields, content-controls, watermark, style resolution), AI diff/apply | No DOM, no PM, no React                |
| `document`      | Editable model: PM schema, plugins, commands, extensions, `Document↔ProseMirror`                                                                                               | JS forever (ProseMirror is JS)         |
| `render-dom`    | Paint `Layout`→DOM, hit-testing, span mapping, scroll, overlay geometry; the canvas measure provider                                                                           | JS forever, **framework-agnostic DOM** |
| `controller`    | Headless orchestration: layout loop, incremental measure, hidden PM view, selection/scroll; imperative API + events                                                            | Framework-agnostic JS                  |
| `react` / `vue` | Thin: lifecycle, event wiring, chrome (toolbar/menus/dialogs)                                                                                                                  | Per-framework                          |

`layout-bridge` today **conflates** two concerns: pure `Document→FlowBlocks`
(belongs in `engine`) and DOM hit-testing (belongs in `render-dom`). Splitting
these concerns requires explicit engine, document, and render boundaries.

## The seams (the contracts)

Everything crosses these; nothing reaches around them.

### Seam 1 — `MeasureProvider`

The engine must not know _how_ text is measured.

```ts
interface MeasureProvider {
  measureWidths(reqs: MeasureRequest[]): number[]; // batched, pure
  fontMetrics(fontString: string): FontMetrics; // ascent/descent/lineHeight
  // Phase B (later): shapeRuns(...) -> positioned glyphs
}
type MeasureRequest = {
  text: string;
  font: string;
  letterSpacing: number;
  horizontalScale: number;
};
type FontMetrics = { ascent: number; descent: number; lineHeight: number };
```

The browser supplies a canvas-backed implementation; a headless/server path
supplies an appropriate implementation for its environment. This inversion
makes the layout engine pure and lets us test or swap measurement backends.

> Note: this seam already half-exists as `layout-engine/measure/measureWorkerProtocol.ts`
> (a serializable `MeasureRequestEntry[] → width[]` contract fulfilled by a
> stateless worker). `MeasureProvider` generalizes that contract.

### Seam 2 — the data lingua franca

`Document`, `Layout`, `FlowBlock`, `Measure` are pure serializable types in
`model`. They are serializable payloads with no behavior.

### Seam 3 — `Document ↔ ProseMirror`

`toDocument(pmDoc)` / `toProseDoc(document)` (today's `fromProseDoc` /
`toProseDoc`). PM is the live editable state; `Document` is what the engine
consumes. **The engine never sees ProseMirror.** Must be incremental-friendly
(convert only changed blocks).

### Seam 4 — `Layout → paint`

Engine emits `Layout` (data); `render-dom` draws it. One-way, pure-data
(already roughly enforced as `layout-engine → layout-painter`). Design `Layout`
to optionally carry positioned-glyph data so Phase-B shaping is additive.

### Seam 5 — `FolioHost` (environment capabilities)

```ts
interface FolioHost {
  measureProvider: MeasureProvider;
  loadFont(spec: FontSpec): Promise<FontResource>;
  readFile?(path: string): Promise<Uint8Array>; // absent on the web
  writeFile?(path: string, bytes: Uint8Array): Promise<void>;
  schedule(cb: () => void): void; // raf vs immediate
}
```

Browser and Node runtimes provide different hosts. The engine and controller
stay identical across supported environments.

### Seam 6 — `FolioEditor` (headless API + events; the adapter linchpin)

```ts
interface FolioEditor {
  mount(container: HTMLElement): void;
  loadDocx(bytes: Uint8Array): Promise<void>;
  getDocx(): Promise<Uint8Array>;
  commands: { applyStyle(...): void; insertTable(...): void; /* ... */ };
  on(evt: "selectionChange" | "docChange" | "layoutComplete", cb: () => void): () => void;
  destroy(): void;
}
```

Framework adapters only instantiate this, forward lifecycle/events, and render
chrome. The editor _surface_ (pages) is painted imperatively by `render-dom`.

### Seam 7 — Operation / Edit API (the agentic surface)

A stable, versioned, documented schema of edit operations over `Document`
(insert/replace/redline/fill-template/restructure), applied deterministically by
the engine with tracked-change provenance (built on today's `ai-edits`
apply/snapshot/diff). Agents emit ops as JSON; the engine applies them headless.
Addressing uses stable, simple ids (block-id/para-id), never raw UUIDs.

## Dependency direction (strictly one-way)

```
model  <-  engine  <-  measure-impl
  ^          ^
document(PM) |
  ^          |
render-dom --'        (render-dom uses model + Layout; ideally PM-free)
  ^
controller  ->  engine, document, render-dom, host
  ^
react / vue  ->  controller
```

`engine` depends only on `model` + the `MeasureProvider` interface. Adapters
depend only on the `controller` API. The 91 reach-ins collapse to
`adapter → FolioEditor`.

## The measurement crux

Seam 1 is the linchpin and the hardest part, because of a fidelity fork:

- **Phase A:** provider returns advance _widths_; engine breaks lines; the
  painter still emits text as DOM and lets the browser shape glyphs. Cheap, but
  provider-measured widths can drift from browser rendering.
- **Phase B:** provider returns _positioned glyphs_; painter blits them; the
  browser does no text layout. Fully deterministic, and the only honest path
  for Arabic/RTL/complex-script fidelity (shaping, not just width). Bigger
  painter change.

Design `Layout` to carry optional glyph positions from day one so B is additive.
Fonts must be provisioned identically to `measure` and `paint` (same fallback
chain) or layout drifts — a sub-problem of Seam 5.

## Incremental layout & async transport

Don't re-serialize the whole `Document` per keystroke. Model the
controller↔engine contract as a stateful `LayoutSession` (which holds cached
measures and the previous layout, fed dirty ranges). Make its methods
**async-tolerant** (promise-returning) from the start: in the browser a
worker call is effectively in-process, but another host may run the engine
across a process boundary or off the UI thread. An async interface preserves
that option at little cost to the direct path.
