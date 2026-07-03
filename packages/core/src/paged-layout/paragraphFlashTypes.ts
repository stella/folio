/**
 * Option shapes for `scrollToParaId(paraId, { highlight })`.
 *
 * DOM-free so non-browser consumers can type-import without pulling in
 * paragraph-flash DOM helpers.
 */

/** Customization for the transient paragraph flash applied by `scrollToParaId`. */
export type ParagraphHighlightOptions = {
  /** CSS color used for the transient paragraph flash. Defaults to yellow. */
  color?: string;
  /** How long the flash remains visible before it is removed. Defaults to 1200ms. */
  durationMs?: number;
};

/** Optional reveal behavior for `scrollToParaId`. */
export type ScrollToParaIdOptions = {
  /** Flash rendered paragraph fragments after scrolling to the paragraph. */
  highlight?: ParagraphHighlightOptions;
};
