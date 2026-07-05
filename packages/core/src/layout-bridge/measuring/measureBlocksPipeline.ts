/**
 * Floating-aware block measurement pipeline (adapter-facing surface).
 *
 * Upstream exposes a generic `measureBlocksWithFloats(blocks, contentWidth,
 * measureBlock, pageGeometry)` orchestrator: it pre-scans blocks for floating
 * exclusion zones, groups co-located floats, then walks the blocks calling a
 * caller-supplied `measureBlock` with the active zones + cumulative Y.
 *
 * @packageDocumentation
 * @public
 */

/**
 * Page geometry (CSS px) used to resolve page/margin-relative anchored objects
 * into content-area coordinates — currently the vertical anchor of a top-pinned
 * `topAndBottom` band. Structurally identical to the painter's `PageGeometry`,
 * declared here rather than imported because layout-bridge must not import from
 * layout-painter (arch boundary); both paths resolve to identical positions.
 *
 * @public
 */
export type FloatPageGeometry = {
  pageWidth: number;
  pageHeight: number;
  marginLeft: number;
  marginTop: number;
  marginRight: number;
  marginBottom: number;
  contentWidth: number;
  contentHeight: number;
};

// PORT-BLOCKED: `measureBlocksWithFloats`.
//
// Our fork already implements this floating-measure orchestration, but inlined
// into the concrete `measureBlocks(blocks, contentWidth, marginTop,
// pageGeometry, fieldValues)` in `layout-engine/measure/measureBlocks.ts`
// (which is what our React adapter calls) rather than as upstream's generic
// `measureBlock`-callback form. Porting the callback-based
// `measureBlocksWithFloats` verbatim would:
//   1. Duplicate ~300 LOC of float extraction + grouping logic that already
//      lives (privately, unexported) inside our `measureBlocks` — the
//      `extractFloatingZones` family there is not re-exported, and the repo
//      rules forbid duplicating existing core helpers.
//   2. Depend on `resolveAnchoredObjectVerticalTop` + an
//      `anchoredObjectPosition.PageGeometry`, which our fork replaced with
//      `bandTopContentY` / `isPageFrameRelativeAnchor` (`layout-engine/
//      textBoxFlow.ts`) driven by a divergent `BandPageGeometry`
//      (`{ pageHeight, marginBottom }`) shape.
// A faithful clean port therefore needs a real architectural reconciliation
// between the two geometry models (out of scope here). The `FloatPageGeometry`
// type is ported above; callers on our fork should use the concrete
// `measureBlocks` from `layout-engine/measure` for the orchestration itself.
