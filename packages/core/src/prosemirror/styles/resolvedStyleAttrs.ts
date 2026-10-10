import { panic } from "better-result";

import {
  effectiveParagraphNumbering,
  effectiveParagraphNumberingReference,
} from "../numberingAttr";
import {
  computeListRendering,
  numberingLevelHasMarkerSlot,
  type NumberingMap,
} from "../../docx/numberingParser";
import { tableOfContentsStyleLevel } from "../../utils/tableOfContentsStyle";
import { paragraphNumberingReference } from "../../docx/numberingReference";
import { setAutospacingBaseValue } from "../autospacingBase";
import { CLEARED_LIST_RENDERING_ATTRS } from "../listMarker";
import { styleResolvedParagraphFormatting } from "../paragraphFormattingProvenance";
import { listRenderingAttrPatch } from "../listRenderingAttrs";
import { paragraphNumberingAttr } from "../numberingAttr";
import {
  paragraphIndentationAttrPatch,
  directParagraphIndentation,
  withDirectParagraphIndentation,
  paragraphIndentationFromFormatting,
  type DirectParagraphIndentation,
} from "../paragraphIndentation";
import type { ParagraphFormatting } from "../../types/document";
import type { ParagraphAttrs, ParagraphAttrsPatch } from "../schema/nodes";
import type { ResolvedParagraphStyle } from "./styleResolver";

/**
 * Shared helper for projecting a resolved paragraph style onto ProseMirror
 * paragraph node attrs.
 *
 * Both `applyStyle` (toolbar style picker) and the Enter handler's
 * next-style switch need to write the same set of style-controlled attrs.
 * Keeping the projection in one place ensures the two paths stay in sync —
 * a style applied via the picker and a style applied on Enter produce
 * identical paragraph attrs.
 */
type ResolvedStyleIdentity = {
  styleId: string;
  styleName?: string;
};

/**
 * The paragraph attrs a style definition controls. Applying a style resets
 * every one of these to the style's value (or `null` to clear), so a prior
 * style's properties (e.g. a heading's spacing) never leak through. Returns
 * a partial attrs object to merge over the paragraph's existing attrs.
 */
export function paragraphAttrsFromResolvedStyle(
  resolved: ResolvedParagraphStyle,
  identity: ResolvedStyleIdentity,
): ParagraphAttrsPatch {
  const ppr = resolved.paragraphFormatting;
  const runFormatting = resolved.runFormatting;
  const hasRunFormatting = !!runFormatting && Object.keys(runFormatting).length > 0;

  return {
    alignment: ppr?.alignment ?? null,
    alignmentFromStyle: ppr?.alignment,
    spaceBefore: ppr?.spaceBefore ?? null,
    spaceAfter: ppr?.spaceAfter ?? null,
    lineSpacing: ppr?.lineSpacing ?? null,
    lineSpacingRule: ppr?.lineSpacingRule ?? null,
    lineSpacingExplicit: null,
    snapToGrid: ppr?.snapToGrid ?? null,
    indentLeft: ppr?.indentLeft ?? null,
    indentRight: ppr?.indentRight ?? null,
    indentFirstLine: ppr?.indentFirstLine ?? null,
    hangingIndent: ppr?.hangingIndent ?? null,
    contextualSpacing: ppr?.contextualSpacing ?? null,
    keepNext: ppr?.keepNext ?? null,
    keepLines: ppr?.keepLines ?? null,
    pageBreakBefore: ppr?.pageBreakBefore ?? null,
    outlineLevel: ppr?.outlineLevel ?? null,
    // Custom paragraph styles (callouts, bordered headings) carry their own
    // `w:pBdr`; the picker and the Enter-into-w:next path both need to apply
    // them, while clearing any source paragraph's leftover borders when the
    // new style has none.
    borders: ppr?.borders ?? null,
    // The style's run defaults drive the caret height in an empty paragraph
    // and the formatting typed text inherits (see EmptyParagraphFormatExtension).
    defaultTextFormatting: hasRunFormatting ? runFormatting : null,
    // The save path reads this to keep the newly applied style's own values
    // out of the paragraph's direct `w:pPr`.
    _resolvedFormatting: styleResolvedParagraphFormatting(ppr),
    _styleResolvedFormatting: styleResolvedParagraphFormatting(ppr),
    _autospacingBase: autospacingBaseFromResolvedParagraphFormatting(ppr),
    _tableOfContentsLevel: tableOfContentsStyleLevel(identity) ?? null,
  };
}

function autospacingBaseFromResolvedParagraphFormatting(
  ppr: ResolvedParagraphStyle["paragraphFormatting"],
): ParagraphAttrs["_autospacingBase"] | null {
  const base: NonNullable<ParagraphAttrs["_autospacingBase"]> = {};
  if (ppr?.beforeAutospacing) {
    setAutospacingBaseValue(base, "before", ppr.spaceBefore);
  }
  if (ppr?.afterAutospacing) {
    setAutospacingBaseValue(base, "after", ppr.spaceAfter);
  }

  return Object.keys(base).length > 0 ? base : null;
}

/**
 * The list attrs a style's `w:pPr/w:numPr` controls (numbering reference plus
 * the baked marker-rendering attrs that `toProseDoc` normally derives from
 * `listRendering` at load time). Returns null when the style defines no
 * numbering — applying such a style leaves any existing list attrs alone, so
 * directly-applied (toolbar) lists survive a style switch the way they do in
 * Word, where numbering is not cleared by applying an unnumbered style.
 *
 * When the style does define numbering, the full attr group is reset so a
 * previous list's marker attrs never leak into the new one. Without the
 * numbering definitions the marker template can't be resolved and the painter
 * falls back to a plain decimal marker.
 *
 * The content-derived marker attrs (`listImplicitChildLevelAdvances`,
 * `listMarkerSecondSlotOffsetTwips`) are nulled — they depend on the
 * paragraph's inline LISTNUM fields, which the picker has no view of. A
 * subsequent save + reload re-derives them from the document content.
 */
export function listAttrsFromResolvedStyle(
  resolved: ResolvedParagraphStyle,
  numbering: NumberingMap | null | undefined,
): ParagraphAttrsPatch | null {
  const numPr = resolved.paragraphFormatting?.numPr;
  if (numPr?.kind !== "reference") {
    return null;
  }

  const { numId, ilvl = 0 } = numPr;
  const attrs = { ...listAttrsFromNumbering({ numId, ilvl }, numbering), numPr: null };
  attrs.numPrFromStyle = paragraphNumberingAttr(paragraphNumberingReference({ numId, ilvl }));
  Object.assign(
    attrs,
    listIndentationProvenancePatch({
      direct: undefined,
      styleFormatting: resolved.paragraphFormatting,
      numberingSource: "style",
      numPr: { numId, ilvl },
      numbering,
    }),
  );
  for (const key of ["indentLeft", "indentRight", "indentFirstLine"] as const) {
    if (attrs[key] == null) Reflect.deleteProperty(attrs, key);
  }
  if (
    attrs.indentFirstLine === undefined &&
    resolved.paragraphFormatting?.hangingIndent === undefined
  ) {
    Reflect.deleteProperty(attrs, "hangingIndent");
  }
  return attrs;
}

/** Project one resolved numbering level into the complete editor attr group. */
export function listAttrsFromNumbering(
  numPr: { numId: number; ilvl: number },
  numbering: NumberingMap | null | undefined,
): ParagraphAttrsPatch {
  const targetNumPr = { numId: numPr.numId, ilvl: numPr.ilvl };
  const rendering = numbering ? computeListRendering(targetNumPr, numbering) : null;
  return {
    ...CLEARED_LIST_RENDERING_ATTRS,
    numPr: paragraphNumberingAttr(paragraphNumberingReference(targetNumPr)),
    ...(rendering && listRenderingAttrPatch(rendering)),
  };
}

/**
 * The numbering level's indentation a directly numbered paragraph reads as its
 * own where it states none: the load path folds the level's `w:ind` into the
 * paragraph (a direct `w:ind` wins per group, left vs first line/hanging).
 * Keep this projection private: callers use the paired provenance helper so
 * effective level values cannot be mistaken for authored indentation.
 */
function listLevelIndentAttrPatch(
  stated: DirectParagraphIndentation | undefined,
  numPr: { numId: number; ilvl: number },
  numbering: NumberingMap | null | undefined,
): ParagraphAttrsPatch {
  const level = numbering?.getLevel(numPr.numId, numPr.ilvl);
  if (!level?.pPr) {
    return {};
  }
  const patch: ParagraphAttrsPatch = {};
  if (stated?.indentLeft === undefined && level.pPr.indentLeft !== undefined) {
    patch.indentLeft = level.pPr.indentLeft;
  }
  if (stated?.indentFirstLine === undefined && numberingLevelHasMarkerSlot(level)) {
    if (level.pPr.indentFirstLine !== undefined) {
      patch.indentFirstLine = level.pPr.indentFirstLine;
    }
    if (level.pPr.hangingIndent !== undefined) {
      patch.hangingIndent = level.pPr.hangingIndent;
    }
  }
  return patch;
}

type ListIndentationProvenanceOptions = {
  direct: DirectParagraphIndentation | undefined;
  styleFormatting: ParagraphFormatting | undefined;
  numberingSource: "style" | "paragraph";
  numPr: { numId: number; ilvl: number } | undefined;
  numbering: NumberingMap | null | undefined;
};

/** Project effective indentation and both inheritance layers together. */
export function listIndentationProvenancePatch({
  direct,
  styleFormatting,
  numberingSource,
  numPr,
  numbering,
}: ListIndentationProvenanceOptions): ParagraphAttrsPatch {
  const style = {
    ...styleResolvedParagraphFormatting(styleFormatting),
    ...(styleFormatting?.hangingIndent === undefined
      ? {}
      : { hangingIndent: styleFormatting.hangingIndent }),
  };
  const level =
    numPr === undefined
      ? {}
      : listLevelIndentAttrPatch(
          numberingSource === "style" ? paragraphIndentationFromFormatting(style) : undefined,
          numPr,
          numbering,
        );
  // A zero inherited offset with no style offset to cancel is the absent
  // default. Authored zero remains in the direct cluster above this layer.
  if (level.indentLeft === 0 && style.indentLeft === undefined) {
    Reflect.deleteProperty(level, "indentLeft");
  }
  if (
    level.indentFirstLine === 0 &&
    style.indentFirstLine === undefined &&
    style.hangingIndent === undefined
  ) {
    Reflect.deleteProperty(level, "indentFirstLine");
    Reflect.deleteProperty(level, "hangingIndent");
  }
  const resolved = {
    ...style,
    ...(typeof level.indentLeft === "number" ? { indentLeft: level.indentLeft } : {}),
    ...(typeof level.indentFirstLine === "number"
      ? { indentFirstLine: level.indentFirstLine }
      : {}),
    ...(typeof level.hangingIndent === "boolean" ? { hangingIndent: level.hangingIndent } : {}),
  };
  return {
    ...paragraphIndentationAttrPatch({
      direct,
      inherited: paragraphIndentationFromFormatting(resolved),
    }),
    _resolvedFormatting: Object.keys(resolved).length === 0 ? null : resolved,
    _styleResolvedFormatting: Object.keys(style).length === 0 ? null : style,
  };
}

/**
 * The indentation a paragraph leaving its list keeps: the level's indentation
 * it only read from its numbering goes with the numbering, back to what its
 * style gives, while an indentation of its own stays. The inverse of
 * {@link listLevelIndentAttrPatch}.
 */
export function listLevelIndentRemovalPatch(
  attrs: Readonly<ParagraphAttrs>,
  numbering: NumberingMap | null | undefined,
): ParagraphAttrsPatch {
  if (effectiveParagraphNumbering(attrs)?.kind !== "reference") {
    return {};
  }
  const direct = directParagraphIndentation(attrs);
  return {
    ...listIndentationProvenancePatch({
      direct,
      styleFormatting: attrs._styleResolvedFormatting,
      numberingSource: "paragraph",
      numPr: undefined,
      numbering,
    }),
    _originalFormatting: withDirectParagraphIndentation(attrs._originalFormatting, direct) ?? null,
  };
}

/** Recompute every level-dependent attr when a paragraph changes list level. */
export function listLevelAttrPatch(
  attrs: Pick<
    ParagraphAttrs,
    | "listImplicitChildLevelAdvances"
    | "_originalFormatting"
    | "_resolvedFormatting"
    | "_styleResolvedFormatting"
    | "numPr"
    | "numPrFromStyle"
    | "indentLeft"
    | "indentRight"
    | "indentFirstLine"
    | "hangingIndent"
  >,
  ilvl: number,
  numbering: NumberingMap | null | undefined,
): ParagraphAttrsPatch {
  const reference = effectiveParagraphNumberingReference(attrs);
  if (reference === undefined) {
    panic("Cannot change list level without an effective numbering reference");
  }
  const numPr = { numId: reference.numId, ilvl };
  const direct = directParagraphIndentation(attrs);
  return {
    ...listAttrsFromNumbering(numPr, numbering),
    ...(attrs.numPr?.kind !== "reference" && attrs.numPrFromStyle?.kind === "reference"
      ? { numPr: paragraphNumberingAttr({ kind: "levelOnly", ilvl: numPr.ilvl }) }
      : {}),
    listImplicitChildLevelAdvances: attrs.listImplicitChildLevelAdvances ?? null,
    ...listIndentationProvenancePatch({
      direct,
      styleFormatting: attrs._styleResolvedFormatting,
      numberingSource:
        attrs.numPr?.kind === "reference" || attrs.numPrFromStyle == null ? "paragraph" : "style",
      numPr,
      numbering,
    }),
    _originalFormatting: withDirectParagraphIndentation(attrs._originalFormatting, direct) ?? null,
  };
}
