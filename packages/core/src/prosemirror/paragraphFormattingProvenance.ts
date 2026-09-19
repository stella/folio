/**
 * Which `w:pPr` fields the editor carries as an EFFECTIVE value, and therefore
 * needs a resolved companion to save correctly.
 *
 * `toProseDoc` seeds a paragraph's formatting attrs with the value the style
 * cascade resolves to, because that is what the editor renders with. A save
 * must write only what the paragraph states itself: an inherited value written
 * back as direct `w:pPr` outranks the style it was read from, and a later edit
 * to that style stops reaching the paragraph. `TableCellAttrs._resolvedBorders`
 * is the same device for the same reason.
 */

import type { ParagraphFormatting } from "../types/document";
import type { ParagraphAttrs } from "./schema/nodes";

/**
 * How each `w:pPr` field reaches the saved model from paragraph attrs.
 *
 * - `style-resolved-attr`: the named attr holds the effective value, so the
 *   write-back filters it against `ParagraphAttrs._resolvedFormatting`.
 * - `direct-provenance`: the field has its own provenance channel and reader
 *   (`alignmentFromStyle`, `spacingFromDocDefaults`, `numPrFromStyle`, …).
 * - `derived`: not authored; recomputed from another field on load and save.
 * - `original-only`: no editor attr writes it back; it survives through
 *   `ParagraphAttrs._originalFormatting`.
 */
export type ParagraphFieldWriteBack =
  | { kind: "style-resolved-attr"; attr: keyof ParagraphAttrs }
  | { kind: "direct-provenance" }
  | { kind: "derived" }
  | { kind: "original-only" };

/**
 * Total over `ParagraphFormatting`, so a new `w:pPr` field cannot land without
 * a decision about how a save tells its authored value from its inherited one.
 */
export const PARAGRAPH_FORMATTING_WRITE_BACK = {
  alignment: { kind: "direct-provenance" },
  bidi: { kind: "style-resolved-attr", attr: "direction" },
  kinsoku: { kind: "style-resolved-attr", attr: "kinsoku" },
  overflowPunctuation: { kind: "style-resolved-attr", attr: "overflowPunctuation" },
  spaceBefore: { kind: "direct-provenance" },
  spaceAfter: { kind: "direct-provenance" },
  lineSpacing: { kind: "direct-provenance" },
  lineSpacingRule: { kind: "direct-provenance" },
  snapToGrid: { kind: "style-resolved-attr", attr: "snapToGrid" },
  beforeAutospacing: { kind: "direct-provenance" },
  afterAutospacing: { kind: "direct-provenance" },
  spacingExplicit: { kind: "derived" },
  indentLeft: { kind: "style-resolved-attr", attr: "indentLeft" },
  indentRight: { kind: "style-resolved-attr", attr: "indentRight" },
  indentFirstLine: { kind: "style-resolved-attr", attr: "indentFirstLine" },
  // Reads as the sign of `indentFirstLine`; its node-spec default is `false`
  // rather than `null`, so the attr carries no absent state of its own.
  hangingIndent: { kind: "direct-provenance" },
  borders: { kind: "style-resolved-attr", attr: "borders" },
  shading: { kind: "style-resolved-attr", attr: "shading" },
  tabs: { kind: "style-resolved-attr", attr: "tabs" },
  // Effective attrs that no save path reads: both legs round-trip through
  // `_originalFormatting`, so neither can materialise an inherited value.
  keepNext: { kind: "original-only" },
  keepLines: { kind: "original-only" },
  widowControl: { kind: "style-resolved-attr", attr: "widowControl" },
  pageBreakBefore: { kind: "style-resolved-attr", attr: "pageBreakBefore" },
  contextualSpacing: { kind: "style-resolved-attr", attr: "contextualSpacing" },
  numPr: { kind: "direct-provenance" },
  numPrFromStyle: { kind: "derived" },
  outlineLevel: { kind: "style-resolved-attr", attr: "outlineLevel" },
  styleId: { kind: "direct-provenance" },
  frame: { kind: "original-only" },
  suppressLineNumbers: { kind: "original-only" },
  suppressAutoHyphens: { kind: "style-resolved-attr", attr: "suppressAutoHyphens" },
  runProperties: { kind: "original-only" },
  // `<w:specVanish/>` on the paragraph mark; the attr drives layout only.
  runInWithNext: { kind: "original-only" },
} as const satisfies Record<keyof ParagraphFormatting, ParagraphFieldWriteBack>;

export type StyleResolvedParagraphField = {
  [Field in keyof typeof PARAGRAPH_FORMATTING_WRITE_BACK]: (typeof PARAGRAPH_FORMATTING_WRITE_BACK)[Field] extends {
    kind: "style-resolved-attr";
  }
    ? Field
    : never;
}[keyof typeof PARAGRAPH_FORMATTING_WRITE_BACK];

export const STYLE_RESOLVED_PARAGRAPH_FIELDS = Object.keys(PARAGRAPH_FORMATTING_WRITE_BACK).filter(
  (field) =>
    PARAGRAPH_FORMATTING_WRITE_BACK[field as keyof typeof PARAGRAPH_FORMATTING_WRITE_BACK].kind ===
    "style-resolved-attr",
) as StyleResolvedParagraphField[];

/**
 * The part of a resolved style cascade a save needs to keep out of direct
 * `w:pPr`. Narrowed to the governed fields so an ordinary paragraph, whose
 * cascade supplies only spacing, stores nothing at all.
 */
export const styleResolvedParagraphFormatting = (
  resolved: ParagraphFormatting | undefined,
): ParagraphFormatting | undefined => {
  if (!resolved) {
    return undefined;
  }
  const governed: ParagraphFormatting = {};
  let found = false;
  for (const field of STYLE_RESOLVED_PARAGRAPH_FIELDS) {
    const value = resolved[field];
    if (value === undefined) {
      continue;
    }
    Reflect.set(governed, field, value);
    found = true;
  }
  return found ? governed : undefined;
};
