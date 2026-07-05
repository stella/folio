// Vue mirror of packages/react/src/components/toolbarUtils.ts —
// only the framework-agnostic exports (highlight color map + helpers).
// The React-side `extractFormattingState` etc. depend on React's
// FormattingAction types and aren't portable; consumer plugins that
// need them should import from the React adapter.
//
// PORT NOTE: upstream re-exported these from
// `@eigenpal/docx-editor-core/utils/highlightColors`. folio-core has no
// `utils/highlightColors` module, so the canonical OOXML highlight table is
// inlined here (ported verbatim from the upstream core module). The
// `w:highlight` attribute only accepts these named colors, so keep the table
// in sync if the highlight palette changes.

/** OOXML highlight color name <-> hex mapping (hex keys, no `#`, uppercase). */
export const HIGHLIGHT_HEX_TO_NAME: Record<string, string> = {
  FFFF00: "yellow",
  "00FF00": "green",
  "00FFFF": "cyan",
  FF00FF: "magenta",
  "0000FF": "blue",
  FF0000: "red",
  "00008B": "darkBlue",
  "008080": "darkCyan",
  "008000": "darkGreen",
  "800080": "darkMagenta",
  "8B0000": "darkRed",
  "808000": "darkYellow",
  "808080": "darkGray",
  C0C0C0: "lightGray",
  "000000": "black",
  FFFFFF: "white",
};

export function mapHexToHighlightName(hex: string): string | null {
  const normalized = hex.replace(/^#/, "").toUpperCase();
  return HIGHLIGHT_HEX_TO_NAME[normalized] || null;
}
