/**
 * Z-index stacking order for the Vue editor chrome — mirrors the React
 * adapter's layering so layered UI stays consistent across adapters instead of
 * drifting into ad-hoc per-component numbers.
 *
 * Order, low to high: page content (0) < selection overlay < decoration layer
 * < image overlay < HF inline editor < ruler < dropdown/popover < context
 * menu/modal.
 */
export const Z_INDEX = {
  selectionOverlay: 10,
  decorationLayer: 11,
  imageOverlay: 15,
  hfInlineEditor: 10,
  ruler: 30,
  dropdown: 100,
  contextMenu: 10000,
} as const;
