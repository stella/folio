/**
 * Selection-overlay composable — owns the text-caret blink + selection-rect
 * painter, the cell-selection highlight, and the image-selection sync that
 * re-derives `selectedImage` from the live PM NodeSelection.
 *
 * PORT-BLOCKED (multiple absent dependencies):
 *   - `applyCellSelectionHighlight` — absent from our core. Upstream ships it in
 *     `layout-bridge/cellSelectionHighlight.ts`; our `layout-bridge` barrel does
 *     not expose it (see CORE-API-MAP Table 4, PORT). The multi-cell highlight
 *     branch cannot run without it.
 *   - `findImageElement` — absent from our core `layout-painter` barrel (which
 *     exposes only `renderPages`/`LayoutPainter`/registry); it is a PORT item.
 *     The image-overlay sync (`syncSelectedImageToSelection`) depends on it.
 *   - `../components/imageSelectionTypes` (`ImageSelectionInfo`) — the Vue-package
 *     image-selection type module has not been ported.
 *   - `../styles/zIndex` (`Z_INDEX`) — the Vue-package z-index token module does
 *     not exist in our fork (there is no `styles/` dir).
 *
 * The DOM primitives it also uses DO resolve, but at drifted paths
 * (`getSelectionRectsFromDom`/`getCaretPositionFromDom` →
 * `@stll/folio-core/layout-bridge/dom/clickToPositionDom`; `findBodyPmAnchor` →
 * `@stll/folio-core/layout-bridge/dom/findBodyPmSpans`). With the four blockers
 * above there is no coherent working subset, so the whole composable is blocked.
 *
 * Unblock: PORT `applyCellSelectionHighlight` and `findImageElement` into core
 * (or a core barrel), then port the `imageSelectionTypes` + `styles/zIndex`
 * Vue-package modules, and re-point the DOM-primitive imports to the drifted
 * paths above.
 */

export {};
