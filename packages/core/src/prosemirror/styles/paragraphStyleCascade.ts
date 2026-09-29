/**
 * The paragraph attrs the style cascade supplies, shared by the load path and
 * the live editor.
 *
 * `toProseDoc` seeds every paragraph with the value its style cascade resolves
 * to (docDefaults, the enclosing table style, the paragraph style chain),
 * because that is what the editor renders with. A paragraph an edit creates —
 * a paste, a structural table command, a replacement that joins paragraphs —
 * must read the same way live, or it paints differently until the document is
 * reopened. Both paths project the cascade through {@link paragraphStyleCascadeAttrs}.
 */

import type { StyleEngine, TableCellParagraphSpacingOverlay } from "../../style-engine";
import type { Paragraph, ParagraphFormatting, TextFormatting } from "../../types/document";
import {
  mergeParagraphNumbering,
  paragraphNumberingReferenceId,
} from "../../docx/numberingReference";
import {
  mergeParagraphFormatting,
  mergeParagraphTabStops,
} from "../../utils/paragraphFormattingMerge";
import { mergeTextFormatting } from "../../utils/textFormattingMerge";
import { paragraphNumberingAttr } from "../numberingAttr";
import { directionFromBidi } from "../paragraphDirection";
import { styleResolvedParagraphFormatting } from "../paragraphFormattingProvenance";
import { lineSpacingProvenanceFromSpacing } from "../paragraphSpacing";
import { stripParagraphMarkOnlyFormatting } from "../runStyleFormatting";
import type { ParagraphAttrs } from "../schema/nodes";
import { cascadeStyleTextFormatting } from "./styleToggleCascade";
import type { TableStyleRegion } from "./tableStyleRegions";

/** The resolver surface the paragraph cascade reads. */
export type ParagraphCascadeResolver = Pick<
  StyleEngine,
  | "getStyle"
  | "getDocDefaults"
  | "getDefaultParagraphStyle"
  | "getDefaultCharacterStyle"
  | "getRunStyleOwnProperties"
  | "resolveParagraphStyleInTable"
>;

/**
 * Resolve an embedded character-style reference without importing
 * `docDefaults`. The caller already has the paragraph cascade, including
 * document defaults, and will layer these own properties over it.
 */
type ParagraphDefaultFormattingResolver = Pick<
  StyleEngine,
  | "getStyle"
  | "getDocDefaults"
  | "getDefaultParagraphStyle"
  | "getDefaultCharacterStyle"
  | "getRunStyleOwnProperties"
>;

export function resolveRunFormattingWithoutDefaults(
  formatting: TextFormatting | undefined,
  styleResolver: ParagraphDefaultFormattingResolver | null,
): TextFormatting | undefined {
  if (!formatting || !styleResolver) {
    return formatting;
  }

  const characterStyleFormatting = formatting.styleId
    ? styleResolver.getRunStyleOwnProperties(formatting.styleId)
    : undefined;
  return cascadeStyleTextFormatting([
    { formatting: characterStyleFormatting, type: "style" },
    { formatting, type: "direct" },
  ]).formatting;
}

/** @internal Recompute a paragraph's inherited run defaults from authored package state. */
export function resolveParagraphDefaultTextFormatting(
  styleId: string | undefined,
  formatting: Paragraph["formatting"] | undefined,
  styleResolver: ParagraphDefaultFormattingResolver,
  options: { includeParagraphMarkRunProperties?: boolean } = {},
): TextFormatting | undefined {
  const style = styleId
    ? (styleResolver.getStyle(styleId) ?? styleResolver.getDefaultParagraphStyle())
    : styleResolver.getDefaultParagraphStyle();
  const paragraphStyleRpr = style?.type === "paragraph" ? style.rPr : undefined;
  // The pPr/rPr block describes the paragraph mark only — see the comment on
  // `stripParagraphMarkOnlyFormatting`. We must NOT route this through
  // `resolveTextFormatting` here, because that folds docDefaults back into
  // the run properties and then overwrites the paragraph style's font
  // (e.g. FootnoteText's Times New Roman) with the docDefault Calibri when
  // merged into the cascade below.
  const rawParagraphMarkRpr =
    options.includeParagraphMarkRunProperties === false ? undefined : formatting?.runProperties;
  const paragraphRunProperties = rawParagraphMarkRpr
    ? stripParagraphMarkOnlyFormatting(
        resolveRunFormattingWithoutDefaults(rawParagraphMarkRpr, styleResolver) ?? {},
      )
    : undefined;

  const orderedBodyToggleFormatting = cascadeStyleTextFormatting(
    [
      { formatting: styleResolver.getDocDefaults()?.rPr, type: "defaults" },
      { formatting: paragraphStyleRpr, type: "style" },
      { formatting: styleResolver.getDefaultCharacterStyle()?.rPr, type: "style" },
    ],
    {
      ordinaryFormatting: mergeTextFormatting(
        mergeTextFormatting(
          styleResolver.getDocDefaults()?.rPr,
          styleResolver.getDefaultCharacterStyle()?.rPr,
        ),
        paragraphStyleRpr,
      ),
    },
  );
  const bodyRunDefaults = orderedBodyToggleFormatting.formatting;
  return cascadeStyleTextFormatting(
    [
      { cascade: orderedBodyToggleFormatting, type: "carried" },
      { formatting: paragraphRunProperties, type: "direct" },
    ],
    {
      ordinaryFormatting: mergeTextFormatting(bodyRunDefaults, paragraphRunProperties),
    },
  ).formatting;
}

export type ParagraphStyleCascadeInput = {
  styleId: string | undefined;
  /** What the paragraph states itself; each stated field wins over the cascade. */
  formatting: ParagraphFormatting | undefined;
  styleResolver: ParagraphCascadeResolver;
  /**
   * The enclosing table style's modeled paragraph fields, layered between
   * docDefaults and the paragraph's own style chain (cell paragraphs only).
   */
  tableParagraphOverlay?: TableCellParagraphSpacingOverlay | undefined;
  /** Whether the paragraph mark's own run properties join the run defaults. */
  includeParagraphMarkRunProperties: boolean;
};

export type ParagraphStyleCascade = {
  /** The effective attrs, with the provenance a save needs to keep inherited values out of `w:pPr`. */
  attrs: ParagraphAttrs;
  /** The resolved style cascade the attrs were read from. */
  stylePpr: ParagraphFormatting | undefined;
};

/**
 * Project a paragraph's style cascade onto its editor attrs: each field the
 * paragraph states wins, every other field reads the cascade, and the
 * provenance attrs record which is which.
 */
export function paragraphStyleCascadeAttrs({
  styleId,
  formatting,
  styleResolver,
  tableParagraphOverlay,
  includeParagraphMarkRunProperties,
}: ParagraphStyleCascadeInput): ParagraphStyleCascade {
  const attrs: ParagraphAttrs = {};
  const set = <K extends keyof ParagraphAttrs>(
    key: K,
    val: ParagraphAttrs[K] | undefined,
  ): void => {
    if (val !== undefined) {
      attrs[key] = val;
    }
  };

  const resolved = styleResolver.resolveParagraphStyleInTable(styleId, tableParagraphOverlay);
  const stylePpr = resolved.paragraphFormatting;
  // What the paragraph would render as if it stated nothing of its own,
  // narrowed to the fields a save could otherwise materialise (see
  // ParagraphAttrs._resolvedFormatting).
  const resolvedFormatting = styleResolvedParagraphFormatting(stylePpr);
  if (resolvedFormatting) {
    attrs._resolvedFormatting = resolvedFormatting;
  }

  // Apply style-based values as defaults (inline overrides)
  set("alignment", formatting?.alignment ?? stylePpr?.alignment);
  set("alignmentFromStyle", stylePpr?.alignment);
  set("spaceBefore", formatting?.spaceBefore ?? stylePpr?.spaceBefore);
  set("spaceAfter", formatting?.spaceAfter ?? stylePpr?.spaceAfter);
  set("lineSpacing", formatting?.lineSpacing ?? stylePpr?.lineSpacing);
  set("lineSpacingRule", formatting?.lineSpacingRule ?? stylePpr?.lineSpacingRule);
  set("lineSpacingExplicit", lineSpacingProvenanceFromSpacing(formatting));
  set("snapToGrid", formatting?.snapToGrid ?? stylePpr?.snapToGrid);
  set("spacingExplicit", formatting?.spacingExplicit);
  const paragraphStyle = styleId
    ? (styleResolver.getStyle(styleId) ?? styleResolver.getDefaultParagraphStyle())
    : styleResolver.getDefaultParagraphStyle();
  const docDefaultSpacing = styleResolver.getDocDefaults()?.pPr;
  // This existing provenance attribute covers every resolved style layer:
  // default, named paragraph, and enclosing table styles. The direct
  // `formatting` object still wins per field.
  const spacingFromStyle: NonNullable<ParagraphAttrs["spacingFromImplicitDefaultStyle"]> = {};
  if (formatting?.spaceBefore === undefined && stylePpr?.spaceBefore !== undefined) {
    spacingFromStyle.before = true;
  }
  if (formatting?.spaceAfter === undefined && stylePpr?.spaceAfter !== undefined) {
    spacingFromStyle.after = true;
  }
  if (spacingFromStyle.before || spacingFromStyle.after) {
    attrs.spacingFromImplicitDefaultStyle = spacingFromStyle;
  }
  const spacingFromDocDefaults: NonNullable<ParagraphAttrs["spacingFromDocDefaults"]> = {};
  if (
    formatting?.spaceBefore === undefined &&
    tableParagraphOverlay?.spaceBefore === undefined &&
    paragraphStyle?.pPr?.spaceBefore === undefined &&
    docDefaultSpacing?.spaceBefore !== undefined
  ) {
    spacingFromDocDefaults.before = true;
  }
  if (
    formatting?.spaceAfter === undefined &&
    tableParagraphOverlay?.spaceAfter === undefined &&
    paragraphStyle?.pPr?.spaceAfter === undefined &&
    docDefaultSpacing?.spaceAfter !== undefined
  ) {
    spacingFromDocDefaults.after = true;
  }
  if (spacingFromDocDefaults.before || spacingFromDocDefaults.after) {
    attrs.spacingFromDocDefaults = spacingFromDocDefaults;
  }
  // When the paragraph explicitly removes the style's numbering (direct
  // numId=0 under a numbered style), the reference layout also drops the
  // style's marker-positioning indents. The paragraph keeps only the indents
  // it states itself (#765: a direct left=357 renders indented instead of
  // hanging the first line back to the margin). Outside that case w:ind
  // merges per attribute: a direct left-only indent keeps the style's
  // firstLine.
  const numberingRemoved =
    formatting?.numPr?.kind === "none" &&
    paragraphNumberingReferenceId(stylePpr?.numPr) !== undefined;
  const numberingStyleIndent = numberingRemoved ? undefined : stylePpr;
  const effectiveIndent = mergeParagraphFormatting(numberingStyleIndent, formatting);
  set("indentLeft", effectiveIndent?.indentLeft);
  set("indentRight", formatting?.indentRight ?? stylePpr?.indentRight);
  set("indentFirstLine", effectiveIndent?.indentFirstLine);
  set("hangingIndent", effectiveIndent?.hangingIndent);
  set("borders", formatting?.borders ?? stylePpr?.borders);
  set("shading", formatting?.shading ?? stylePpr?.shading);
  set("tabs", mergeParagraphTabStops(stylePpr?.tabs, formatting?.tabs));
  set("kinsoku", formatting?.kinsoku ?? stylePpr?.kinsoku);
  set("overflowPunctuation", formatting?.overflowPunctuation ?? stylePpr?.overflowPunctuation);
  set("suppressAutoHyphens", formatting?.suppressAutoHyphens ?? stylePpr?.suppressAutoHyphens);

  // Page break control
  set("pageBreakBefore", formatting?.pageBreakBefore ?? stylePpr?.pageBreakBefore);
  set("keepNext", formatting?.keepNext ?? stylePpr?.keepNext);
  set("keepLines", formatting?.keepLines ?? stylePpr?.keepLines);
  set("widowControl", formatting?.widowControl ?? stylePpr?.widowControl);
  set("contextualSpacing", formatting?.contextualSpacing ?? stylePpr?.contextualSpacing);
  // Run-in heading (`<w:specVanish/>` on the paragraph mark) — see
  // ParagraphAttrs.runInWithNext.
  set("runInWithNext", formatting?.runInWithNext ?? stylePpr?.runInWithNext);

  // Outline level (for TOC)
  set("outlineLevel", formatting?.outlineLevel ?? stylePpr?.outlineLevel);

  // Text direction — a direct or style-sourced `w:bidi` is an authoritative
  // manual decision (auto-detection must not override it).
  set("direction", directionFromBidi(formatting?.bidi ?? stylePpr?.bidi));

  set(
    "defaultTextFormatting",
    resolveParagraphDefaultTextFormatting(styleId, formatting, styleResolver, {
      includeParagraphMarkRunProperties,
    }),
  );

  // A direct numPr may carry only ilvl while the style supplies numId.
  // Merge the two fields so the effective list keeps the style's numbering
  // identity. A direct numId (including 0) is authoritative.
  const styleNumbering = stylePpr?.numPr;
  if (
    styleNumbering?.kind === "reference" &&
    (formatting?.numPr === undefined || formatting.numPr.kind === "levelOnly")
  ) {
    const merged = mergeParagraphNumbering(styleNumbering, formatting?.numPr);
    if (merged !== undefined) {
      attrs.numPr = paragraphNumberingAttr(merged);
    }
    attrs.numPrFromStyle = paragraphNumberingAttr(styleNumbering);
  }

  return { attrs, stylePpr };
}

/**
 * Pick the modeled paragraph fields out of a table style's (or conditional
 * region's) `w:pPr` for use as the cell-paragraph cascade overlay.
 */
export function extractTableParagraphOverlay(
  pPr: ParagraphFormatting | undefined,
): TableCellParagraphSpacingOverlay | undefined {
  if (!pPr) {
    return undefined;
  }
  const overlay: TableCellParagraphSpacingOverlay = {};
  if (pPr.spaceBefore !== undefined) {
    overlay.spaceBefore = pPr.spaceBefore;
  }
  if (pPr.spaceAfter !== undefined) {
    overlay.spaceAfter = pPr.spaceAfter;
  }
  if (pPr.lineSpacing !== undefined) {
    overlay.lineSpacing = pPr.lineSpacing;
  }
  if (pPr.lineSpacingRule !== undefined) {
    overlay.lineSpacingRule = pPr.lineSpacingRule;
  }
  if (pPr.contextualSpacing !== undefined) {
    overlay.contextualSpacing = pPr.contextualSpacing;
  }
  if (pPr.frame !== undefined) {
    overlay.frame = pPr.frame;
  }
  return Object.keys(overlay).length > 0 ? overlay : undefined;
}

/**
 * The paragraph overlay a table style gives one cell's paragraphs: its base
 * `w:pPr`, then each applicable region's (`regions`, lowest precedence first,
 * from `tableCellStyleRegions`). A table naming no style, or a style the part
 * does not define, reads the default table style (ECMA-376 §17.7.6).
 */
export function tableCellParagraphOverlay(
  styleResolver: Pick<StyleEngine, "getStyle" | "getDefaultTableStyle">,
  tableStyleId: string | null | undefined,
  regions: readonly TableStyleRegion[],
): TableCellParagraphSpacingOverlay | undefined {
  const style =
    (tableStyleId ? styleResolver.getStyle(tableStyleId) : undefined) ??
    styleResolver.getDefaultTableStyle();
  if (!style) {
    return undefined;
  }
  let overlay: ParagraphFormatting | undefined = extractTableParagraphOverlay(style.pPr);
  for (const region of regions) {
    const conditional = style.tblStylePr?.find(({ type }) => type === region);
    overlay = mergeParagraphFormatting(overlay, extractTableParagraphOverlay(conditional?.pPr));
  }
  return overlay;
}
