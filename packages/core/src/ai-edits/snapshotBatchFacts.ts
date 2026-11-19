import { panic } from "better-result";
import type { Node as PMNode, ResolvedPos } from "prosemirror-model";
import { TableMap } from "prosemirror-tables";

import { paragraphEndsItsContainer } from "../prosemirror/containerFinalParagraph";
import type { BatchClaim } from "./batch-claims";
import { holdsOnlyDeletedContent } from "./pendingJoin";
import { isEmptyTableCell } from "./table-cell-mutations";
import type {
  TableMutationPlanTarget,
  TableRectangle,
} from "./table-mutation-plan";
import type {
  TableColumnDeletion,
  TableRowDeletion,
} from "./table-row-column-mutations";

/** Required snapshot facts shared by the native and canonical batch planners. */
export type SnapshotBatchOperation =
  | Exclude<
      BatchClaim,
      { type: "deleteBlock" | "tableRemoval" | "tableStructure" | "unclaimed" }
    >
  | { type: "deleteBlock"; block: number; end: number; node: PMNode }
  | { type: "deleteTable"; position: number; node: PMNode }
  | {
      type: "deleteTableRow";
      deletion: TableRowDeletion;
      row: PMNode;
      table: PMNode;
    }
  | { type: "deleteTableColumn"; deletion: TableColumnDeletion }
  | {
      type: "tableStructure";
      target: Exclude<TableMutationPlanTarget, { type: "none" }>;
      axis: "row" | "column" | "cell";
    };

/** A paragraph mark that is a pending insertion (not a relocation's end). */
export const isInsertedPPrMark = (value: unknown): boolean =>
  typeof value === "object" &&
  value !== null &&
  "kind" in value &&
  value.kind === "ins";

/** A paragraph mark that is a plain pending deletion (not a relocation's source). */
export const isPlainDeletedPPrMark = (value: unknown): boolean =>
  typeof value === "object" &&
  value !== null &&
  "kind" in value &&
  value.kind === "del";

/**
 * Whether everything after `position` in its container is paragraphs pending
 * deletion: a join there, once accepted, has no words after it, so a merge's
 * separator would separate nothing.
 */
export const followedOnlyByDeletedParagraphs = (
  doc: PMNode,
  position: number,
): boolean => {
  const at = doc.resolve(position);
  let deletes = false;
  for (let index = at.index(); index < at.parent.childCount; index++) {
    const node = at.parent.child(index);
    if (node.type.name !== "paragraph" || !holdsOnlyDeletedContent(node))
      return false;
    deletes ||=
      node.content.size > 0 || isPlainDeletedPPrMark(node.attrs["pPrMark"]);
  }
  return deletes;
};

/** The cells a column deletion removes: every cell standing in that column alone. */
const deletedColumnCellRanges = (
  doc: PMNode,
  { tablePosition, columnIndex }: TableColumnDeletion,
): { from: number; to: number }[] => {
  const table = doc.nodeAt(tablePosition);
  if (!table || table.type.spec["tableRole"] !== "table") {
    return [];
  }
  const map = TableMap.get(table);
  const tableStart = tablePosition + 1;
  const ranges: { from: number; to: number }[] = [];
  const seen = new Set<number>();
  for (let row = 0; row < map.height; row++) {
    const relative = map.map[row * map.width + columnIndex];
    if (relative === undefined || seen.has(relative)) {
      continue;
    }
    seen.add(relative);
    const rectangle = map.findCell(relative);
    const cell = table.nodeAt(relative);
    if (
      !cell ||
      rectangle.left !== columnIndex ||
      rectangle.right !== columnIndex + 1
    ) {
      continue;
    }
    ranges.push({
      from: tableStart + relative,
      to: tableStart + relative + cell.nodeSize,
    });
  }
  return ranges;
};

type DeletedRowAffectedRangesOptions = {
  doc: PMNode;
  deletion: TableRowDeletion;
  row: PMNode;
};

/** A row deletion also changes cells whose vertical merge began above it. */
const deletedRowAffectedRanges = ({
  doc,
  deletion,
  row,
}: DeletedRowAffectedRangesOptions): { from: number; to: number }[] => {
  const ranges = [
    { from: deletion.rowPosition, to: deletion.rowPosition + row.nodeSize },
  ];
  const table = doc.nodeAt(deletion.tablePosition);
  if (table?.type.spec["tableRole"] !== "table") {
    return panic("A resolved row deletion lost its table", {
      position: deletion.tablePosition,
    });
  }
  const map = TableMap.get(table);
  const seen = new Set<number>();
  for (let column = 0; column < map.width; column++) {
    const relative = map.map[deletion.rowIndex * map.width + column];
    if (relative === undefined || seen.has(relative)) continue;
    seen.add(relative);
    const position = deletion.tableStart + relative;
    if (position >= deletion.rowPosition) continue;
    const cell = table.nodeAt(relative);
    if (cell) ranges.push({ from: position, to: position + cell.nodeSize });
  }
  return ranges;
};

/**
 * The blank cells a merge folds into the first cell of its rectangle. Only a
 * blank cell can fold away tracked, so an edit that fills one decides whether
 * the merge applies in one mode and not the other; see `batch-claims.ts`. A
 * cell with content folds directly and is refused tracked whatever the batch
 * does inside it.
 */
const foldedBlankCellRanges = (
  doc: PMNode,
  {
    tablePosition,
    rectangle,
  }: { tablePosition: number; rectangle: TableRectangle },
): { from: number; to: number }[] => {
  const table = doc.nodeAt(tablePosition);
  if (table?.type.spec["tableRole"] !== "table") {
    return panic("A resolved cell merge lost its table", {
      position: tablePosition,
    });
  }
  const map = TableMap.get(table);
  const tableStart = tablePosition + 1;
  const kept = map.map[rectangle.top * map.width + rectangle.left];
  return map.cellsInRect(rectangle).flatMap((relative) => {
    const cell = relative === kept ? null : table.nodeAt(relative);
    return cell && isEmptyTableCell(cell)
      ? [
          {
            from: tableStart + relative,
            to: tableStart + relative + cell.nodeSize,
          },
        ]
      : [];
  });
};

/**
 * The paragraphs before `paragraph`, which ends its container, nearest first,
 * up to the container's start or another block, past those that go whole, as
 * `finalParagraphPredecessor` looks past them. See the `deleteBlock` claim's
 * `paragraphsBefore`.
 */
const paragraphsBeforeFinal = (
  $paragraph: ResolvedPos,
  paragraph: PMNode,
): { block: number; keepsBreak: boolean }[] => {
  const paragraphs: { block: number; keepsBreak: boolean }[] = [];
  let position = $paragraph.pos;
  for (let index = $paragraph.index() - 1; index >= 0; index--) {
    const node = $paragraph.parent.child(index);
    position -= node.nodeSize;
    if (node.type !== paragraph.type) break;
    const mark: unknown = node.attrs["pPrMark"];
    if (isPlainDeletedPPrMark(mark) && holdsOnlyDeletedContent(node)) continue;
    paragraphs.push({ block: position, keepsBreak: mark == null });
  }
  return paragraphs;
};

type SnapshotBatchClaimOptions = {
  item: SnapshotBatchOperation;
  doc: PMNode;
  producesTrackedChanges: boolean;
};

/** Derive claims from the unchanged document against which the batch was read. */
export const snapshotBatchClaim = ({
  item,
  doc,
  producesTrackedChanges,
}: SnapshotBatchClaimOptions): Exclude<BatchClaim, { type: "unclaimed" }> => {
  switch (item.type) {
    case "text":
    case "annotation":
    case "split":
    case "rewriteBlock":
    case "paragraphProperties":
    case "merge":
    case "insertion":
      return item;
    case "deleteBlock": {
      const at = doc.resolve(item.block);
      const endsItsContainer = paragraphEndsItsContainer(
        at,
        item.node.type.name,
      );
      const keepsParagraph = producesTrackedChanges
        ? endsItsContainer
        : endsItsContainer && at.nodeBefore?.type.name !== item.node.type.name;
      return {
        type: "deleteBlock",
        block: item.block,
        end: item.end,
        keepsParagraph,
        endsContainer: endsItsContainer,
        paragraphsBefore:
          producesTrackedChanges && endsItsContainer
            ? paragraphsBeforeFinal(at, item.node)
            : [],
      };
    }
    case "deleteTable":
      return {
        type: "tableRemoval",
        table: item.position,
        wholeTable: true,
        axis: "table",
        ranges: [
          { from: item.position, to: item.position + item.node.nodeSize },
        ],
      };
    case "deleteTableRow":
      return {
        type: "tableRemoval",
        table: item.deletion.tablePosition,
        wholeTable: item.table.childCount === 1,
        axis: "row",
        ranges:
          item.table.childCount === 1
            ? [
                {
                  from: item.deletion.tablePosition,
                  to: item.deletion.tablePosition + item.table.nodeSize,
                },
              ]
            : deletedRowAffectedRanges({
                doc,
                deletion: item.deletion,
                row: item.row,
              }),
      };
    case "deleteTableColumn":
      return {
        type: "tableRemoval",
        table: item.deletion.tablePosition,
        wholeTable: false,
        axis: "column",
        ranges: deletedColumnCellRanges(doc, item.deletion),
      };
    case "tableStructure":
      return {
        type: "tableStructure",
        table: item.target.tablePosition,
        axis: item.axis,
        folded:
          item.target.type === "mergeCells"
            ? foldedBlankCellRanges(doc, item.target)
            : [],
      };
  }
};
