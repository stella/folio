/**
 * Resolve the style cascade onto paragraphs an edit creates.
 *
 * A loaded paragraph carries its effective formatting as attrs: `toProseDoc`
 * projects the style cascade (docDefaults, the paragraph style chain) onto
 * every paragraph, and the painter reads those attrs. A paragraph an edit
 * creates — a paste, a structural table command, a replacement that joins
 * paragraphs, a cut that empties the document — starts from the node spec's
 * defaults instead, so it paints without its inherited spacing and run
 * defaults until the document is saved and reopened.
 *
 * {@link resolveEditedParagraphStyles} closes that gap in one place: after
 * every edit it projects the same cascade onto each paragraph the edit touched
 * and that the load path never saw, filling only what the paragraph leaves
 * unstated, with the provenance a save needs to keep those inherited values
 * out of the paragraph's own `w:pPr`.
 */

import type { Node as PMNode, ResolvedPos } from "prosemirror-model";
import { isHistoryTransaction } from "prosemirror-history";
import type { EditorState, Transaction } from "prosemirror-state";
import { TableMap } from "prosemirror-tables";
import { Mapping } from "prosemirror-transform";

import type { NumberingMap } from "../../docx/numberingParser";
import { resolveTableLook } from "../../docx/tableLook";
import type { StyleEngine } from "../../style-engine";
import type { ParagraphFormatting } from "../../types/document";
import { tableOfContentsStyleLevel } from "../../utils/tableOfContentsStyle";
import {
  expectParagraphAttrs,
  expectTableAttrs,
  expectTableCellAttrs,
  expectTableRowAttrs,
} from "../attrs";
import { setAutospacingBaseValue } from "../autospacingBase";
import { directionToAuthoredBidi } from "../paragraphDirection";
import type { ParagraphAttrs } from "../schema/nodes";
import {
  type ParagraphCascadeResolver,
  paragraphStyleCascadeAttrs,
  tableCellParagraphOverlay,
} from "../styles/paragraphStyleCascade";
import {
  type TableStyleRegion,
  tableCellStyleRegions,
  tableRowBand,
} from "../styles/tableStyleRegions";
import type { TableCellParagraphSpacingOverlay } from "../styles/styleResolver";
import { listAttrsFromResolvedStyle } from "../styles/resolvedStyleAttrs";

/** The ranges of `doc` the given transactions changed, in its own positions. */
const changedRanges = (transactions: readonly Transaction[]): Array<[number, number]> => {
  const ranges: Array<[number, number]> = [];
  const maps = transactions.flatMap((transaction) => transaction.mapping.maps);
  let offset = 0;
  for (const transaction of transactions) {
    const own = transaction.mapping.maps.length;
    if (transaction.docChanged && !isHistoryTransaction(transaction)) {
      for (const [index, map] of transaction.mapping.maps.entries()) {
        const following = new Mapping(maps.slice(offset + index + 1));
        // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror StepMap.forEach
        map.forEach((_oldFrom, _oldTo, from, to) => {
          // oxlint-disable-next-line unicorn/no-array-method-this-argument -- ProseMirror Mapping.map(pos, assoc)
          ranges.push([following.map(from, -1), following.map(to, 1)]);
        });
      }
    }
    offset += own;
  }
  return ranges;
};

/**
 * What the paragraph states itself, as the cascade reads it: every attr the
 * node holds a value for. Only called for paragraphs the load path never
 * projected, whose attrs hold nothing inherited yet.
 */
const statedFormatting = (attrs: ParagraphAttrs): ParagraphFormatting => {
  const formatting: ParagraphFormatting = {};
  const copy = <Key extends keyof ParagraphFormatting & keyof ParagraphAttrs>(key: Key): void => {
    const value = attrs[key];
    if (value !== null && value !== undefined) {
      Reflect.set(formatting, key, value);
    }
  };
  copy("alignment");
  copy("spaceBefore");
  copy("spaceAfter");
  copy("lineSpacing");
  copy("lineSpacingRule");
  copy("snapToGrid");
  copy("indentLeft");
  copy("indentRight");
  copy("indentFirstLine");
  if (attrs.indentFirstLine !== null && attrs.indentFirstLine !== undefined) {
    copy("hangingIndent");
  }
  copy("borders");
  copy("shading");
  copy("tabs");
  copy("kinsoku");
  copy("overflowPunctuation");
  copy("suppressAutoHyphens");
  copy("pageBreakBefore");
  copy("keepNext");
  copy("keepLines");
  copy("widowControl");
  copy("contextualSpacing");
  copy("runInWithNext");
  copy("outlineLevel");
  copy("numPr");
  const bidi = directionToAuthoredBidi(attrs.direction);
  if (bidi !== undefined) {
    formatting.bidi = bidi;
  }
  return formatting;
};

type ParagraphStyleResolver = ParagraphCascadeResolver & Pick<StyleEngine, "getDefaultTableStyle">;

const isUnset = (value: unknown): boolean => value === null || value === undefined;

/**
 * The attrs the cascade supplies that `paragraph` leaves unset, or null when
 * it already reads the way a reopened document would.
 */
const cascadePatch = (
  node: PMNode,
  resolver: ParagraphStyleResolver,
  tableParagraphOverlay: TableCellParagraphSpacingOverlay | undefined,
  numbering: () => NumberingMap | null,
): Partial<ParagraphAttrs> | null => {
  const attrs = expectParagraphAttrs(node);
  const styleId = attrs.styleId ?? undefined;
  const formatting = statedFormatting(attrs);
  const { attrs: cascade, stylePpr } = paragraphStyleCascadeAttrs({
    styleId,
    formatting,
    styleResolver: resolver,
    tableParagraphOverlay,
    includeParagraphMarkRunProperties: false,
  });
  const patch: Partial<ParagraphAttrs> = {};
  const fill = <Key extends keyof ParagraphAttrs>(key: Key): void => {
    const value = cascade[key];
    if (!isUnset(value) && isUnset(attrs[key])) {
      patch[key] = value;
    }
  };

  // A side the paragraph marks as its own stays as stated, even when empty.
  const lineProvenance = attrs.lineSpacingExplicit;
  const filledSides = {
    before: !attrs.spacingExplicit?.before && isUnset(attrs.spaceBefore),
    after: !attrs.spacingExplicit?.after && isUnset(attrs.spaceAfter),
  };
  if (filledSides.before) {
    fill("spaceBefore");
  }
  if (filledSides.after) {
    fill("spaceAfter");
  }
  if (lineProvenance !== "value" && lineProvenance !== "both" && lineProvenance !== true) {
    fill("lineSpacing");
  }
  if (lineProvenance !== "rule" && lineProvenance !== "both" && lineProvenance !== true) {
    fill("lineSpacingRule");
  }
  for (const key of ["spacingFromDocDefaults", "spacingFromImplicitDefaultStyle"] as const) {
    const inherited = cascade[key];
    const side = {
      ...(inherited?.before && filledSides.before ? { before: true } : {}),
      ...(inherited?.after && filledSides.after ? { after: true } : {}),
    };
    if (side.before || side.after) {
      patch[key] = { ...attrs[key], ...side };
    }
  }
  if (isUnset(attrs.indentFirstLine) && attrs.hangingIndent !== true) {
    fill("indentFirstLine");
    if (cascade.hangingIndent === true) {
      patch.hangingIndent = true;
    }
  }
  for (const key of [
    "alignment",
    "alignmentFromStyle",
    "_resolvedFormatting",
    "snapToGrid",
    "indentLeft",
    "indentRight",
    "borders",
    "shading",
    "tabs",
    "kinsoku",
    "overflowPunctuation",
    "suppressAutoHyphens",
    "pageBreakBefore",
    "keepNext",
    "keepLines",
    "widowControl",
    "contextualSpacing",
    "runInWithNext",
    "outlineLevel",
    "direction",
    "defaultTextFormatting",
  ] as const) {
    fill(key);
  }
  if (styleId !== undefined && isUnset(attrs._tableOfContentsLevel)) {
    const styleName = resolver.getStyle(styleId)?.name;
    const level = tableOfContentsStyleLevel({ styleId, ...(styleName ? { styleName } : {}) });
    if (level !== undefined) {
      patch._tableOfContentsLevel = level;
    }
  }
  if (
    stylePpr?.numPr?.kind === "reference" &&
    isUnset(attrs.numPr) &&
    isUnset(attrs.numPrFromStyle)
  ) {
    const listAttrs = listAttrsFromResolvedStyle({ paragraphFormatting: stylePpr }, numbering());
    // The numbering level's indents apply only where the paragraph and its
    // style state none; the rest of the list attr group is the style's own.
    const takesLevelIndent: Record<string, boolean> = {
      indentLeft: isUnset(attrs.indentLeft) && patch.indentLeft === undefined,
      indentFirstLine:
        isUnset(attrs.indentFirstLine) &&
        attrs.hangingIndent !== true &&
        patch.indentFirstLine === undefined,
    };
    takesLevelIndent["hangingIndent"] = takesLevelIndent["indentFirstLine"] === true;
    for (const [key, value] of Object.entries(listAttrs ?? {})) {
      if (takesLevelIndent[key] === false) {
        continue;
      }
      Reflect.set(patch, key, value);
    }
  }
  const autospacingBase: NonNullable<ParagraphAttrs["_autospacingBase"]> = {};
  if (stylePpr?.beforeAutospacing && patch.spaceBefore !== undefined) {
    setAutospacingBaseValue(autospacingBase, "before", patch.spaceBefore);
  }
  if (stylePpr?.afterAutospacing && patch.spaceAfter !== undefined) {
    setAutospacingBaseValue(autospacingBase, "after", patch.spaceAfter);
  }
  if (Object.keys(autospacingBase).length > 0) {
    patch._autospacingBase = { ...attrs._autospacingBase, ...autospacingBase };
  }
  return Object.keys(patch).length > 0 ? patch : null;
};

/** The table-style regions for the cell at `depth` of `$pos`, as the load path picks them. */
const cellStyleRegions = ($pos: ResolvedPos, depth: number): TableStyleRegion[] => {
  const table = $pos.node(depth - 2);
  const row = $pos.node(depth - 1);
  const map = TableMap.get(table);
  const cellPos = $pos.before(depth) - $pos.start(depth - 2);
  const rect = map.findCell(cellPos);
  const look = resolveTableLook(expectTableAttrs(table).look ?? undefined);
  return tableCellStyleRegions({
    look,
    rowIndex: rect.top,
    totalRows: map.height,
    column: rect.left,
    colspan: rect.right - rect.left,
    totalColumns: map.width,
    rowBand: tableRowBand(look, rect.top, map.height),
    rowConditionalFormat: expectTableRowAttrs(row)._originalFormatting?.conditionalFormat,
    cellConditionalFormat: expectTableCellAttrs($pos.node(depth))._originalFormatting
      ?.conditionalFormat,
  });
};

/**
 * The table-style paragraph overlay for a paragraph directly inside a table
 * cell (or a block container in one), as the load path layers it.
 */
const cellParagraphOverlay = (
  $pos: ResolvedPos,
  resolver: ParagraphStyleResolver,
): TableCellParagraphSpacingOverlay | undefined => {
  for (let depth = $pos.depth; depth > 0; depth -= 1) {
    const ancestor = $pos.node(depth);
    const role: unknown = ancestor.type.spec["tableRole"];
    if (role === "cell" || role === "header_cell") {
      const table = $pos.node(depth - 2);
      const styleId: unknown = table.attrs["styleId"];
      return tableCellParagraphOverlay(
        resolver,
        typeof styleId === "string" ? styleId : null,
        cellStyleRegions($pos, depth),
      );
    }
    if (ancestor.type.name === "textBox") {
      return undefined;
    }
  }
  return undefined;
};

/**
 * Project the style cascade onto every paragraph the transactions touched
 * that has not been projected yet. Returns the transaction to append, or null
 * when every touched paragraph already reads as a reopened document would.
 */
export const resolveEditedParagraphStyles = (
  transactions: readonly Transaction[],
  state: EditorState,
  resolver: ParagraphStyleResolver,
  numbering: () => NumberingMap | null,
): Transaction | null => {
  const ranges = changedRanges(transactions);
  if (ranges.length === 0) {
    return null;
  }
  const tr = state.tr;
  const seen = new Set<number>();
  const docSize = state.doc.content.size;
  for (const [rangeFrom, rangeTo] of ranges) {
    const from = Math.max(0, Math.min(rangeFrom, docSize));
    const to = Math.max(from, Math.min(rangeTo, docSize));
    state.doc.nodesBetween(from, to, (node, pos) => {
      if (node.type.name !== "paragraph") {
        return !node.isTextblock;
      }
      if (seen.has(pos)) {
        return false;
      }
      seen.add(pos);
      // A paragraph the load path projected carries the source record it was
      // read from; its unset attrs are an edit's decision, not a gap.
      if (node.attrs["_originalFormatting"] !== null) {
        return false;
      }
      const patch = cascadePatch(
        node,
        resolver,
        cellParagraphOverlay(state.doc.resolve(pos), resolver),
        numbering,
      );
      if (patch) {
        for (const [key, value] of Object.entries(patch)) {
          const current: unknown = node.attrs[key];
          if (current !== value && !(isUnset(current) && isUnset(value))) {
            tr.setNodeAttribute(pos, key, value);
          }
        }
      }
      return false;
    });
  }
  // Attribute steps map no positions, so suggesting mode, which marks the
  // ranges a step inserts, never reads this as an authored change.
  return tr.docChanged ? tr : null;
};
