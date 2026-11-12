// A field-totality gate in a test file is read by no compiler, so the sample
// map the run property order property draws from lives here, where the
// package typecheck compiles it.
import type { RecordField } from "../../model/document";
import type { TextFormatting } from "../../model/formatting";

/** One minimal `TextFormatting` per model field, each stating that field. */
export const RUN_PROPERTY_SAMPLES = {
  bold: { bold: true },
  boldCs: { boldCs: false },
  italic: { italic: true },
  italicCs: { italicCs: false },
  underline: { underline: { style: "single" } },
  strike: { strike: true },
  doubleStrike: { doubleStrike: true },
  vertAlign: { vertAlign: "superscript" },
  smallCaps: { smallCaps: true },
  allCaps: { allCaps: true },
  hidden: { hidden: true },
  noProof: { noProof: true },
  color: { color: { rgb: "1F4E79" } },
  highlight: { highlight: "yellow" },
  shading: { shading: { pattern: "clear", fill: { rgb: "D9D9D9" } } },
  fontSize: { fontSize: 24 },
  fontSizeCs: { fontSizeCs: 26 },
  fontFamily: {
    fontFamily: { ascii: "Georgia", hAnsi: "Georgia", cs: "Amiri" },
  },
  language: { language: { val: "cs-CZ" } },
  spacing: { spacing: 20 },
  position: { position: 6 },
  scale: { scale: 110 },
  kerning: { kerning: 18 },
  effect: { effect: "shimmer" },
  emphasisMark: { emphasisMark: "dot" },
  emboss: { emboss: true },
  imprint: { imprint: true },
  outline: { outline: true },
  shadow: { shadow: true },
  rtl: { rtl: true },
  cs: { cs: true },
  styleId: { styleId: "Emphasis" },
  // Keeps this map total; the compiled writer has no capture channel, so this
  // sample contributes no child to its emitted `w:rPr`.
  preserved: { preserved: { children: [{ index: 0, xml: "<w:bdr/>" }] } },
} as const satisfies Record<RecordField<TextFormatting>, TextFormatting>;
