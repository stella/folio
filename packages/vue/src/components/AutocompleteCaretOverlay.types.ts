export type AutocompleteCaretRect = {
  /** Overlay-relative pixel x. */
  x: number;
  /** Overlay-relative pixel y (top of the cursor line). */
  y: number;
  /** Line height in pixels at the anchor. */
  lineHeight: number;
  /** Available width from the caret to the page's right content edge. */
  maxWidth?: number | undefined;
};

export type AutocompleteCaretOverlayProps = {
  /** Position of the autocomplete anchor, or null while idle. */
  caret: AutocompleteCaretRect | null;
  /** The full streamed ghost text so far. */
  text: string;
  /** Whether tokens are still arriving. */
  isStreaming: boolean;
};
