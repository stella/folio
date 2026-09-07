/**
 * A table's own properties, read and matched across two documents.
 *
 * `w:tblPr`, `w:trPr` and `w:tcPr` are the part of a table no block carries: a
 * paragraph snapshot says which cell it sits in and nothing about the cell's
 * width, span, shading, borders or margins, nor about the row's height or
 * header flag, nor about the table's own width, indent, justification,
 * borders, cell margins, look or style. A comparison that reproduces every
 * word and none of that hands back a table that is not the one it was compared
 * to.
 *
 * Two things happen here. {@link projectTableGeometry} renders those
 * properties as text, so the round-trip self-check sees a lost property the
 * same way it sees a lost paragraph. {@link matchTableGeometry} moves them:
 * where a paired table, row or cell differs from the one it pairs with, the
 * target's properties are written and the previous set is recorded as
 * `w:tblPrChange` / `w:trPrChange` / `w:tcPrChange`, which is what a reject
 * restores.
 *
 * The scope of each is not a hand-kept list. It is exactly the attrs a reject
 * of the matching change element restores, read off the reject patches
 * themselves — so a property that reject can restore is a property the
 * comparison carries and checks, and the three can never drift apart.
 */

import type { Node as PMNode } from "prosemirror-model";
import type { Transaction } from "prosemirror-state";

import { markStructuralChange } from "../prosemirror/extensions/features/ParagraphChangeTrackerExtension";

import {
  tableCellRejectAttrPatch,
  tableRejectAttrPatch,
  tableRowRejectAttrPatch,
} from "../prosemirror/commands/propertyChangeScope";
import {
  tableAttrsToFormatting,
  tableCellAttrsToFormatting,
  tableRowAttrsToFormatting,
} from "../prosemirror/conversion/fromProseDoc";
import type { TableCellAttrs } from "../prosemirror/schema/nodes";
import type { TableCellFormatting, TableFormatting, TableRowFormatting } from "../types/document";
import { canonicalJson } from "../utils/canonicalJson";
import type { FolioStoryTable } from "./snapshot";

const ORIGINAL_FORMATTING = "_originalFormatting";

/**
 * The formatting a parser stored verbatim for the save path. It is deliberately
 * outside every projection below: a document saved by an editor materializes
 * style-resolved properties into its own `w:tcPr`, so two packages that render
 * the same table can store different formatting for it, and comparing what was
 * stored would report a difference no redline can or should represent. What is
 * compared is the effective set the properties resolve to.
 */
const withoutOriginalFormatting = (keys: readonly string[]): string[] =>
  keys.filter((key) => key !== ORIGINAL_FORMATTING);

/**
 * Attrs a `w:tblPrChange` reject restores, and therefore the attrs a match
 * writes and the projection reads. `columnWidths` is deliberately outside
 * them: the grid is `w:tblGrid`, not `w:tblPr`, and the editable model has no
 * `w:tblGridChange` to record a change of it against.
 */
const TABLE_SCOPED_ATTRS = withoutOriginalFormatting(Object.keys(tableRejectAttrPatch(undefined)));
const ROW_SCOPED_ATTRS = withoutOriginalFormatting(Object.keys(tableRowRejectAttrPatch(undefined)));
const CELL_SCOPED_ATTRS = withoutOriginalFormatting(
  Object.keys(tableCellRejectAttrPatch(undefined, undefined)),
);

/**
 * `false` and absent are the same thing for every property in scope here: they
 * are all presence flags (`w:tblHeader`, `w:hidden`, `w:noWrap`), and a parser
 * that materializes an absent one as `false` must not read as a difference
 * from one that leaves it unset.
 */
const scopedAttrs = (
  attrs: Record<string, unknown>,
  keys: readonly string[],
): Record<string, unknown> => {
  const scoped: Record<string, unknown> = {};
  for (const key of keys) {
    const value = attrs[key];
    scoped[key] = value === false ? null : (value ?? null);
  }
  return scoped;
};

const cellProjection = (cell: PMNode): string =>
  canonicalJson({
    colspan: cell.attrs["colspan"],
    rowspan: cell.attrs["rowspan"],
    ...scopedAttrs(cell.attrs, CELL_SCOPED_ATTRS),
  });

const rowProjection = (row: PMNode): string => {
  const cells: string[] = [];
  row.forEach((cell) => {
    cells.push(cellProjection(cell));
  });
  return `${canonicalJson(scopedAttrs(row.attrs, ROW_SCOPED_ATTRS))}|${cells.join("|")}`;
};

/**
 * One line per table, in the document order tables are numbered in: the
 * table's own properties, then each row's and each cell's. Nested tables get
 * their own lines rather than being folded into their parent's, so a
 * difference names the table it is in.
 */
export const projectTableGeometry = (tables: readonly FolioStoryTable[]): string[] =>
  tables.map(({ node }) => {
    const rows: string[] = [];
    node.forEach((row) => {
      rows.push(rowProjection(row));
    });
    return `${canonicalJson(scopedAttrs(node.attrs, TABLE_SCOPED_ATTRS))}#${rows.join("#")}`;
  });

/** Where a cell sits: which table, which row of it, which cell of that row. */
export type TableCellCoordinate = {
  tableIndex: number;
  rowIndex: number;
  cellIndex: number;
};

/** One base cell and the target cell it was aligned with. */
export type TableGeometryPairing = {
  base: TableCellCoordinate;
  target: TableCellCoordinate;
};

type MatchTableGeometryOptions = {
  tr: Transaction;
  /** The base story's tables, with the positions the transaction will write at. */
  baseTables: readonly FolioStoryTable[];
  /** The target story's tables, by the same index the pairings name. */
  targetTables: ReadonlyMap<number, PMNode>;
  pairings: readonly TableGeometryPairing[];
  revision: { author: string; date: string; idSeed: number };
};

export type MatchTableGeometryResult = {
  /** First revision id a following batch may allocate. */
  nextRevisionId: number;
  /** Tables, rows and cells whose properties the match moved. */
  matched: number;
};

type PropertyChangeTarget = {
  position: number;
  attrs: Record<string, unknown>;
  changeAttr: "tblPrChange" | "trPrChange" | "tcPrChange";
  changeType: "tablePropertyChange" | "tableRowPropertyChange" | "tableCellPropertyChange";
  previousFormatting: unknown;
};

const childPositions = (node: PMNode, start: number): number[] => {
  const positions: number[] = [];
  let offset = start + 1;
  node.forEach((child) => {
    positions.push(offset);
    offset += child.nodeSize;
  });
  return positions;
};

type PropertyScope<TFormatting> = {
  keys: readonly string[];
  /**
   * The property set a node's attrs serialize to — what its `w:tblPr` /
   * `w:trPr` / `w:tcPr` would say. It is the record a change element has to
   * store, because it holds the EFFECTIVE properties: some of a cell's come
   * from the table or a table style rather than from its own element, and a
   * record built from what the parser stored would not restore them.
   */
  formattingOf: (node: PMNode) => TFormatting | undefined;
  /** What a reject of this change element restores, from a stored property set. */
  rejectPatch: (
    previousFormatting: TFormatting | undefined,
    liveFormatting: TFormatting | undefined,
  ) => Record<string, unknown>;
  changeAttr: PropertyChangeTarget["changeAttr"];
  changeType: PropertyChangeTarget["changeType"];
};

/**
 * `Node["attrs"]` is an open record, and the converters below want the node
 * type's own shape. The two span counts are the only members the schema always
 * carries a value for, so naming them is what turns the record into one.
 */
const cellAttrsOf = (node: PMNode): TableCellAttrs => ({
  ...effectiveAttrs(node),
  colspan: typeof node.attrs["colspan"] === "number" ? node.attrs["colspan"] : 1,
  rowspan: typeof node.attrs["rowspan"] === "number" ? node.attrs["rowspan"] : 1,
});

/**
 * The node's attrs with the style cascade's own values cleared, so the
 * converter treats every effective value as one the node states.
 *
 * A change element stores the COMPLETE previous property set and a reject
 * rebuilds the live properties from it alone, so the record has to hold what
 * the node renders with — including a border its table style supplied. The
 * save path wants the opposite (write only what the node states, or the
 * inherited value becomes an override), which is what the resolved companions
 * are for; a record is the one place they get in the way.
 */
const effectiveAttrs = (node: PMNode): Record<string, unknown> => ({
  ...node.attrs,
  _resolvedBorders: null,
  _resolvedMargins: null,
  _resolvedCellMargins: null,
});

const TABLE_SCOPE = {
  keys: TABLE_SCOPED_ATTRS,
  formattingOf: (node) => tableAttrsToFormatting(effectiveAttrs(node)),
  rejectPatch: (previousFormatting) => tableRejectAttrPatch(previousFormatting),
  changeAttr: "tblPrChange",
  changeType: "tablePropertyChange",
} as const satisfies PropertyScope<TableFormatting>;

const ROW_SCOPE = {
  keys: ROW_SCOPED_ATTRS,
  formattingOf: (node) => tableRowAttrsToFormatting(node.attrs),
  rejectPatch: (previousFormatting) => tableRowRejectAttrPatch(previousFormatting),
  changeAttr: "trPrChange",
  changeType: "tableRowPropertyChange",
} as const satisfies PropertyScope<TableRowFormatting>;

const CELL_SCOPE = {
  keys: CELL_SCOPED_ATTRS,
  formattingOf: (node) => tableCellAttrsToFormatting(cellAttrsOf(node)),
  rejectPatch: (previousFormatting, liveFormatting) =>
    tableCellRejectAttrPatch(previousFormatting, liveFormatting),
  changeAttr: "tcPrChange",
  changeType: "tableCellPropertyChange",
} as const satisfies PropertyScope<TableCellFormatting>;

const scopedValues = (attrs: Record<string, unknown>, keys: readonly string[]): string =>
  canonicalJson(scopedAttrs(attrs, keys));

/**
 * One node's properties, matched to the node it was paired with, or `null`
 * when the difference cannot be recorded.
 *
 * A change element stores the COMPLETE previous property set, and rejecting it
 * rebuilds the live properties from that record alone. So the record is built
 * from what the node would SERIALIZE, not from what its parser stored: a cell
 * inherits borders and margins from the table and from a table style, and a
 * record that omitted them would reject to a third document.
 *
 * The check is direct, and it is what keeps the round trip exact: rebuild the
 * live values from the record and see whether they come back. Where they do
 * not, the difference is left alone rather than written as a revision that
 * cannot be undone.
 */
const propertyChangeFor = <TFormatting>(
  base: PMNode,
  target: PMNode,
  position: number,
  scope: PropertyScope<TFormatting>,
): PropertyChangeTarget | null => {
  const live = scopedValues(base.attrs, scope.keys);
  if (live === scopedValues(target.attrs, scope.keys)) {
    return null;
  }
  const previousFormatting = scope.formattingOf(base);
  if (
    scopedValues(scope.rejectPatch(previousFormatting, previousFormatting), scope.keys) !== live
  ) {
    return null;
  }
  return {
    position,
    // The target's effective values, and the formatting THEY serialize to, so
    // the two stay consistent: a document saved by an editor materializes
    // style-resolved properties into its own element, and only the effective
    // set is comparable across two packages.
    attrs: {
      ...scopedAttrs(target.attrs, scope.keys),
      [ORIGINAL_FORMATTING]: scope.formattingOf(target) ?? null,
    },
    changeAttr: scope.changeAttr,
    changeType: scope.changeType,
    previousFormatting,
  };
};

/**
 * Copy a paired table's, row's and cell's properties from the document it was
 * compared against, recording the previous set so a reject restores it.
 *
 * Pairings name cells, because that is what the block alignment resolves; the
 * row and the table each cell sits in are matched with it. A property set that
 * already agrees is left alone, so an unchanged table produces no revision.
 *
 * `colspan` / `rowspan` never move: they shape the table map, and changing one
 * without restructuring the rows around it leaves the map inconsistent with
 * its own grid. That is the rule a reject of a `w:tcPrChange` follows too.
 */
export const matchTableGeometry = ({
  tr,
  baseTables,
  targetTables,
  pairings,
  revision,
}: MatchTableGeometryOptions): MatchTableGeometryResult => {
  const baseByIndex = new Map(baseTables.map((table) => [table.index, table]));
  const targets: PropertyChangeTarget[] = [];
  const claimed = new Set<number>();

  const consider = <TFormatting>(
    base: PMNode,
    target: PMNode,
    position: number,
    scope: PropertyScope<TFormatting>,
  ): void => {
    if (claimed.has(position)) {
      return;
    }
    claimed.add(position);
    const change = propertyChangeFor(base, target, position, scope);
    if (change) {
      targets.push(change);
    }
  };

  for (const { base, target } of pairings) {
    const baseTable = baseByIndex.get(base.tableIndex);
    const targetTable = targetTables.get(target.tableIndex);
    if (!baseTable || !targetTable) {
      continue;
    }
    const baseRow = baseTable.node.maybeChild(base.rowIndex);
    const targetRow = targetTable.maybeChild(target.rowIndex);
    const rowPosition = childPositions(baseTable.node, baseTable.start)[base.rowIndex];
    if (!baseRow || !targetRow || rowPosition === undefined) {
      continue;
    }
    const baseCell = baseRow.maybeChild(base.cellIndex);
    const targetCell = targetRow.maybeChild(target.cellIndex);
    const cellPosition = childPositions(baseRow, rowPosition)[base.cellIndex];
    if (!baseCell || !targetCell || cellPosition === undefined) {
      continue;
    }
    consider(baseTable.node, targetTable, baseTable.start, TABLE_SCOPE);
    consider(baseRow, targetRow, rowPosition, ROW_SCOPE);
    consider(baseCell, targetCell, cellPosition, CELL_SCOPE);
  }

  let revisionId = revision.idSeed;
  for (const target of targets) {
    const node = tr.doc.nodeAt(target.position);
    if (!node) {
      continue;
    }
    tr.setNodeMarkup(target.position, undefined, {
      ...node.attrs,
      ...target.attrs,
      [target.changeAttr]: [
        {
          type: target.changeType,
          info: { id: revisionId, author: revision.author, date: revision.date },
          ...(target.previousFormatting != null && {
            previousFormatting: target.previousFormatting,
          }),
        },
      ],
    });
    revisionId += 1;
  }
  if (revisionId > revision.idSeed) {
    // A property change writes no paragraph, and the incremental save reuses
    // a part whose paragraphs did not change. Saying the table's structure
    // moved is what makes the save rewrite it.
    markStructuralChange(tr);
  }
  return { nextRevisionId: revisionId, matched: targets.length };
};
