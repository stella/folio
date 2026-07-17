/**
 * Property-change reject scope + patch builders.
 *
 * Word tracks formatting changes by storing the COMPLETE old property set
 * inside the change element (`w:pPrChange` stores the old `CT_PPrBase`,
 * `w:sectPrChange` the old `CT_SectPrBase`, `w:tblPrChange`/`w:trPrChange`/
 * `w:tcPrChange` the old tblPr/trPr/tcPr). Rejecting such a change therefore
 * replaces the live properties WHOLESALE within that scope: a property the
 * change ADDED (present on the live node, absent from the stored old set)
 * must reset, not survive. Properties OUTSIDE the stored scope (an inline
 * sectPr on a paragraph, header/footer references inside a sectPr, the
 * paragraph-mark rPr) are separately modeled and must survive a reject.
 *
 * This module is the single place that maps stored "previous" records onto
 * ProseMirror node attrs. Records come in two shapes:
 *
 * - parser-shaped: `parseParagraphPropertyChanges` (docx/paragraphParser.ts)
 *   stores a `ParagraphFormatting`, whose keys mostly match paragraph attr
 *   names except `bidi` (attr `direction`) and the autospacing flags (attr
 *   `_autospacingBase`); table records store `TableFormatting` /
 *   `TableRowFormatting` / `TableCellFormatting`.
 * - attr-shaped: editor-created suggestions (ListExtension) snapshot the
 *   paragraph attrs directly (including list-rendering bookkeeping attrs).
 *
 * Both shapes normalize here; per key, an attr-shaped value wins when present.
 */

import type {
  BorderSpec,
  CellMargins,
  ParagraphFormatting,
  SectionProperties,
  TableCellFormatting,
  TableFormatting,
  TableRowFormatting,
} from "../../types/document";
import { setAutospacingBaseValue } from "../autospacingBase";
import { directionFromBidi } from "../paragraphDirection";
import type { ParagraphAttrs } from "../schema/nodes";

type AttrPatch = Record<string, unknown>;

/**
 * Paragraph attrs governed by a `w:pPrChange` — the attrs
 * `paragraphFormattingToAttrs` (conversion/toProseDoc.ts) can produce from a
 * parsed `CT_PPrBase` (ECMA-376 §17.13.5.29). Rejecting a pPrChange restores
 * the stored old pPr wholesale for exactly these keys: a key absent from the
 * stored record resets to the schema default (`null`).
 *
 * Deliberately OUT of scope (preserved across a reject):
 * - identity/structure: `paraId`, `textId`, `sectionBreakType`,
 *   `_sectionProperties`, `_propertyChanges`, `pPrMark`, `bookmarks`,
 *   `_emptyHyperlinks`, `renderedPageBreakBefore`
 * - paragraph-mark run properties — `pPr/rPr` is CT_PPr-only, not part of the
 *   CT_PPrBase payload a pPrChange stores: `defaultTextFormatting`,
 *   `runInWithNext` (`w:specVanish` lives in that rPr)
 * - load-time style/numbering bookkeeping the command layer cannot recompute
 *   without a style resolver: `numPrFromStyle`, the `list*` rendering attrs,
 *   `spacingFromDocDefaults`, `spacingFromImplicitDefaultStyle`
 */
export const PPR_CHANGE_SCOPED_ATTR_KEYS = [
  "styleId",
  "numPr",
  "alignment",
  "spaceBefore",
  "spaceAfter",
  "lineSpacing",
  "lineSpacingRule",
  "snapToGrid",
  "spacingExplicit",
  "indentLeft",
  "indentRight",
  "indentFirstLine",
  "hangingIndent",
  "borders",
  "shading",
  "tabs",
  "pageBreakBefore",
  "keepNext",
  "keepLines",
  "widowControl",
  "contextualSpacing",
  "outlineLevel",
  "direction",
  "_autospacingBase",
] as const satisfies readonly (keyof ParagraphAttrs)[];

const PPR_CHANGE_SCOPED_ATTR_KEY_SET: ReadonlySet<string> = new Set(PPR_CHANGE_SCOPED_ATTR_KEYS);

/**
 * Parser-shaped `ParagraphFormatting` keys that must NOT merge through to
 * attrs verbatim: they either normalize to a differently named attr
 * (`bidi` → `direction`, autospacing flags → `_autospacingBase`), have no
 * attr representation and round-trip via `_originalFormatting` only
 * (`frame`, `suppressLineNumbers`, `suppressAutoHyphens`), or sit outside
 * the CT_PPrBase scope entirely (`runProperties`, `runInWithNext` — the
 * paragraph-mark rPr) and so must never overwrite the live value on reject.
 */
const PPR_PARSER_ONLY_KEYS: ReadonlySet<string> = new Set([
  "bidi",
  "beforeAutospacing",
  "afterAutospacing",
  "runProperties",
  "runInWithNext",
  "frame",
  "suppressLineNumbers",
  "suppressAutoHyphens",
  "numPrFromStyle",
]);

/**
 * `ParagraphFormatting` keys inside the CT_PPrBase scope, used to rebuild
 * `_originalFormatting` after a reject (old pPr wholesale within scope,
 * paragraph-mark rPr fields preserved from the live original).
 */
const PPR_CHANGE_SCOPED_FORMATTING_KEYS = [
  "styleId",
  "numPr",
  "alignment",
  "bidi",
  "spaceBefore",
  "spaceAfter",
  "lineSpacing",
  "lineSpacingRule",
  "beforeAutospacing",
  "afterAutospacing",
  "spacingExplicit",
  "indentLeft",
  "indentRight",
  "indentFirstLine",
  "hangingIndent",
  "borders",
  "shading",
  "tabs",
  "pageBreakBefore",
  "keepNext",
  "keepLines",
  "widowControl",
  "contextualSpacing",
  "outlineLevel",
  "frame",
  "suppressLineNumbers",
  "suppressAutoHyphens",
] as const satisfies readonly (keyof ParagraphFormatting)[];

/**
 * Build the attr patch that rejecting one pPrChange applies to a paragraph:
 * every in-scope key set to the stored previous value, or reset to `null`
 * when the stored old pPr does not carry it. Keys the record captured beyond
 * the scoped set (editor-created list suggestions snapshot list-rendering
 * bookkeeping attrs) merge through 1:1 so their pre-change values restore too.
 */
export function paragraphRejectAttrPatch(
  previousFormatting: Record<string, unknown> | null | undefined,
): AttrPatch {
  const prev = previousFormatting ?? {};
  const patch: AttrPatch = {};
  for (const key of PPR_CHANGE_SCOPED_ATTR_KEYS) {
    patch[key] = Object.hasOwn(prev, key) ? (prev[key] ?? null) : null;
  }
  // Parser-shaped fallbacks for the two renamed attrs (attr-shaped records
  // carrying `direction` / `_autospacingBase` already won above).
  if (!Object.hasOwn(prev, "direction")) {
    patch["direction"] = directionFromBidi(prev["bidi"] as boolean | null | undefined);
  }
  if (!Object.hasOwn(prev, "_autospacingBase")) {
    patch["_autospacingBase"] = autospacingBaseFromFormatting(prev);
  }
  for (const [key, value] of Object.entries(prev)) {
    if (PPR_CHANGE_SCOPED_ATTR_KEY_SET.has(key) || PPR_PARSER_ONLY_KEYS.has(key)) {
      continue;
    }
    patch[key] = value ?? null;
  }
  return patch;
}

/**
 * Rebuild the paragraph's `_originalFormatting` (the serializer's pPr source)
 * after a reject: the stored old pPr wholesale within CT_PPrBase scope, with
 * the paragraph-mark rPr fields (`runProperties`, `runInWithNext`) preserved
 * from the live original — a pPrChange cannot store them, so a reject must
 * not drop them.
 */
export function paragraphRejectOriginalFormatting(
  previousFormatting: Record<string, unknown> | null | undefined,
  liveOriginal: unknown,
): ParagraphFormatting | null {
  const prev = previousFormatting ?? {};
  const result: Record<string, unknown> = {};
  for (const key of PPR_CHANGE_SCOPED_FORMATTING_KEYS) {
    const value = Object.hasOwn(prev, key) ? prev[key] : undefined;
    if (value != null) {
      result[key] = value;
    }
  }
  const live = isRecord(liveOriginal) ? liveOriginal : {};
  if (live["runProperties"] != null) {
    result["runProperties"] = live["runProperties"];
  }
  if (live["runInWithNext"] != null) {
    result["runInWithNext"] = live["runInWithNext"];
  }
  if (Object.keys(result).length === 0) {
    return null;
  }
  // SAFETY: every key written above is a ParagraphFormatting key; values come
  // from a stored ParagraphFormatting (or an attr-shaped snapshot whose
  // overlapping keys share the same value shapes).
  return result as ParagraphFormatting;
}

function autospacingBaseFromFormatting(prev: Record<string, unknown>): AttrPatch | null {
  const before = prev["beforeAutospacing"] === true;
  const after = prev["afterAutospacing"] === true;
  if (!before && !after) {
    return null;
  }
  const base: NonNullable<ParagraphAttrs["_autospacingBase"]> = {};
  if (before) {
    setAutospacingBaseValue(base, "before", prev["spaceBefore"]);
  }
  if (after) {
    setAutospacingBaseValue(base, "after", prev["spaceAfter"]);
  }
  return base;
}

/**
 * Rejected section properties: the stored old sectPr wholesale, preserving
 * the live header/footer references — `EG_HdrFtrReferences` is not part of
 * the `CT_SectPrBase` payload a `w:sectPrChange` stores (ECMA-376
 * §17.13.5.32), so those children survive a reject. The caller re-attaches
 * whatever `propertyChanges` remain unresolved.
 */
export function sectionRejectProperties(
  live: SectionProperties,
  previousProperties: SectionProperties | undefined,
): SectionProperties {
  const restored: SectionProperties = { ...previousProperties };
  delete restored.propertyChanges;
  if (live.headerReferences) {
    restored.headerReferences = live.headerReferences;
  } else {
    delete restored.headerReferences;
  }
  if (live.footerReferences) {
    restored.footerReferences = live.footerReferences;
  } else {
    delete restored.footerReferences;
  }
  return restored;
}

/**
 * Attr patch for rejecting a `w:tblPrChange`: the stored old tblPr wholesale
 * (`w:tblPrChange` stores the complete previous tblPr, so every tblPr-derived
 * attr resets when absent). `columnWidths` comes from `w:tblGrid`, not tblPr,
 * and is untouched.
 */
export function tableRejectAttrPatch(previousFormatting: TableFormatting | undefined): AttrPatch {
  return {
    styleId: previousFormatting?.styleId ?? null,
    width: previousFormatting?.width?.value ?? null,
    widthType: previousFormatting?.width?.type ?? null,
    justification: previousFormatting?.justification ?? null,
    floating: previousFormatting?.floating ?? null,
    cellMargins: previousFormatting?.cellMargins
      ? cellMarginsToAttr(previousFormatting.cellMargins)
      : null,
    look: previousFormatting?.look ?? null,
    borders: previousFormatting?.borders ?? null,
    _originalFormatting: previousFormatting ?? null,
  };
}

/** Attr patch for rejecting a `w:trPrChange`: the stored old trPr wholesale. */
export function tableRowRejectAttrPatch(
  previousFormatting: TableRowFormatting | undefined,
): AttrPatch {
  return {
    height: previousFormatting?.height?.value ?? null,
    heightRule: previousFormatting?.heightRule ?? null,
    isHeader: previousFormatting?.header ?? false,
    hidden: previousFormatting?.hidden ?? null,
    _originalFormatting: previousFormatting ?? null,
  };
}

/**
 * Attr patch for rejecting a `w:tcPrChange`: the stored old tcPr wholesale
 * for the non-structural attrs. `gridSpan` / `vMerge` are structural in
 * ProseMirror (`colspan` / `rowspan` shape the table grid); restoring them
 * from the record would desync the grid without a full table restructure, so
 * the live structure is kept and the rebuilt `_originalFormatting` inherits
 * the live original's `gridSpan` / `vMerge` to stay consistent with it.
 * (Word would also restore tracked merges; folio deliberately does not.)
 */
export function tableCellRejectAttrPatch(
  previousFormatting: TableCellFormatting | undefined,
  liveOriginalFormatting: TableCellFormatting | null | undefined,
): AttrPatch {
  const restoredOriginal: TableCellFormatting = { ...previousFormatting };
  delete restoredOriginal.gridSpan;
  delete restoredOriginal.vMerge;
  if (liveOriginalFormatting?.gridSpan !== undefined) {
    restoredOriginal.gridSpan = liveOriginalFormatting.gridSpan;
  }
  if (liveOriginalFormatting?.vMerge !== undefined) {
    restoredOriginal.vMerge = liveOriginalFormatting.vMerge;
  }
  return {
    width: previousFormatting?.width?.value ?? null,
    widthType: previousFormatting?.width?.type ?? null,
    verticalAlign: previousFormatting?.verticalAlign ?? null,
    backgroundColor: previousFormatting?.shading?.fill?.rgb ?? null,
    textDirection: previousFormatting?.textDirection ?? null,
    noWrap: previousFormatting?.noWrap ?? null,
    borders: previousFormatting?.borders ? cellBordersToAttr(previousFormatting.borders) : null,
    margins: previousFormatting?.margins ? cellMarginsToAttr(previousFormatting.margins) : null,
    _originalFormatting: Object.keys(restoredOriginal).length > 0 ? restoredOriginal : null,
  };
}

type SideMargins = { top?: number; bottom?: number; left?: number; right?: number };

function cellMarginsToAttr(margins: CellMargins): SideMargins {
  const result: SideMargins = {};
  if (margins.top?.value !== undefined) {
    result.top = margins.top.value;
  }
  if (margins.bottom?.value !== undefined) {
    result.bottom = margins.bottom.value;
  }
  if (margins.left?.value !== undefined) {
    result.left = margins.left.value;
  }
  if (margins.right?.value !== undefined) {
    result.right = margins.right.value;
  }
  return result;
}

type SideBorders = { top?: BorderSpec; bottom?: BorderSpec; left?: BorderSpec; right?: BorderSpec };

function cellBordersToAttr(borders: NonNullable<TableCellFormatting["borders"]>): SideBorders {
  const result: SideBorders = {};
  if (borders.top) {
    result.top = borders.top;
  }
  if (borders.bottom) {
    result.bottom = borders.bottom;
  }
  if (borders.left) {
    result.left = borders.left;
  }
  if (borders.right) {
    result.right = borders.right;
  }
  return result;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}
