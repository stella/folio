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
 * Paragraph records use the canonical authored-property state. Table records
 * retain their table-specific formatting models below.
 */

import type { Node as PMNode } from "prosemirror-model";
import { isParagraphFormattingPropertyKey } from "@stll/docx-core/model";

import type {
  BorderSpec,
  CellMargins,
  SectionProperties,
  TableCellFormatting,
  TableFormatting,
  TableRowFormatting,
} from "../../types/document";

import { expectParagraphAttrs } from "../attrs";
import {
  isGovernedParagraphAttr,
  transitionParagraphProperties,
  type ParagraphPropertyProjection,
} from "../paragraphPropertyMutation";
import {
  authoredParagraphPropertiesFromFormatting,
  expectParagraphPropertyState,
  type ParagraphPropertyProjectionContext,
} from "../paragraphPropertyState";
import type { ParagraphAttrs, ParagraphPropertyChangeAttrs } from "../schema/nodes";

/** Editor-only suggestions are rebased away before save; every other entry emits `w:pPrChange`. */
export const hasSerializableParagraphPropertyChange = (
  changes: readonly ParagraphPropertyChangeAttrs[] | null | undefined,
): boolean => changes?.some(({ info }) => info.provenance !== "suggested") === true;

type AttrPatch = Record<string, unknown>;
type ParagraphPropertySnapshot = NonNullable<ParagraphPropertyChangeAttrs["previousFormatting"]>;

export type ParagraphPropertyChangeRemoval =
  | { type: "unchanged" }
  | {
      type: "keep-live";
      remaining: ParagraphPropertyChangeAttrs[];
    }
  | {
      type: "restore-previous";
      remaining: ParagraphPropertyChangeAttrs[];
      previousFormatting: ParagraphPropertyChangeAttrs["previousFormatting"];
    };

/**
 * Remove selected entries from an oldest-to-newest paragraph-property chain.
 * A retained entry immediately after a removed run inherits that run's
 * earliest previous snapshot. Only a removed trailing run restores the live
 * paragraph properties. This makes repeated scoped rejection independent of
 * the order in which revisions are resolved.
 */
export const removeParagraphPropertyChanges = (
  changes: readonly ParagraphPropertyChangeAttrs[],
  shouldRemove: (change: ParagraphPropertyChangeAttrs) => boolean,
): ParagraphPropertyChangeRemoval => {
  const remaining: ParagraphPropertyChangeAttrs[] = [];
  let removedPrevious: ParagraphPropertyChangeAttrs["previousFormatting"];
  let removedAny = false;
  let removingRun = false;

  for (const change of changes) {
    if (shouldRemove(change)) {
      if (!removingRun) {
        removedPrevious = change.previousFormatting;
        removingRun = true;
      }
      removedAny = true;
      continue;
    }
    if (!removingRun) {
      remaining.push(change);
      continue;
    }

    const rebased = { ...change };
    if (removedPrevious === undefined) {
      Reflect.deleteProperty(rebased, "previousFormatting");
    } else {
      rebased.previousFormatting = removedPrevious;
    }
    remaining.push(rebased);
    removedPrevious = undefined;
    removingRun = false;
  }

  if (!removedAny) {
    return { type: "unchanged" };
  }
  if (!removingRun) {
    return { type: "keep-live", remaining };
  }
  return { type: "restore-previous", remaining, previousFormatting: removedPrevious };
};

/** The in-scope paragraph properties as they stand, for a `w:pPrChange` record. */
export const paragraphPropertiesSnapshot = (node: PMNode): ParagraphPropertySnapshot => {
  const attrs = expectParagraphAttrs(node);
  return structuredClone(
    expectParagraphPropertyState(attrs._paragraphPropertyState).authoredPPr,
  );
};

type ParagraphPropertyRejectionProjectionOptions = {
  attrs: ParagraphAttrs;
  previousFormatting: ParagraphPropertySnapshot | null | undefined;
  context?: ParagraphPropertyProjectionContext;
};

/** Canonical accepted/rejected-view projection for a removed `w:pPrChange`. */
export const projectParagraphPropertyRejection = ({
  attrs,
  previousFormatting,
  context,
}: ParagraphPropertyRejectionProjectionOptions): ParagraphPropertyProjection => {
  const previous = previousFormatting ?? {};
  const nextAttrs = { ...attrs };
  for (const [key, value] of Object.entries(previous)) {
    if (isParagraphFormattingPropertyKey(key) || isGovernedParagraphAttr(key)) {
      continue;
    }
    Reflect.set(nextAttrs, key, value ?? null);
  }
  return transitionParagraphProperties({
    attrs: nextAttrs,
    state: {
      type: "update",
      authored: {
        type: "replace",
        authoredPPr: authoredParagraphPropertiesFromFormatting(previous),
      },
      context: context === undefined ? { type: "preserve" } : { type: "replace", context },
    },
  });
};

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
