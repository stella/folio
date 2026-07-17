import { Result } from "better-result";
import { Fragment, type Mark, type Node as PMNode, type Schema } from "prosemirror-model";
import type { EditorState, Transaction } from "prosemirror-state";
import {
  columnIsHeader,
  removeColumn,
  removeRow,
  rowIsHeader,
  TableMap,
  tableNodeTypes,
} from "prosemirror-tables";

import { expectRunPropertyChangeMarkAttrs, expectTableCellAttrs } from "../prosemirror/attrs";
import {
  marksToTextFormatting,
  standaloneTableCellFromProseMirror,
} from "../prosemirror/conversion/fromProseDoc";
import { standaloneTableCellToProseMirror } from "../prosemirror/conversion/toProseDoc";
import { markStructuralChange } from "../prosemirror/extensions/features/ParagraphChangeTrackerExtension";
import { getFolioParaIdFromBlockId } from "../types/block-id";
import type { RunPropertyChange, TableCell, TableCellFormatting } from "../types/document";
import { buildCleanBlockText } from "./clean-text";
import {
  hasInlineEmphasis,
  parseInlineEmphasisRuns,
  stripInlineEmphasisMarkers,
} from "./inline-emphasis";
import { hashFolioAIBlockText, normalizeFolioAIBlockText } from "./snapshot";
import type {
  FolioAIEditAppliedOperation,
  FolioAIEditApplyMode,
  FolioAIEditApplyResult,
  FolioAIEditOperation,
  FolioAIEditSnapshot,
  FolioAIEditSkipReason,
  FolioAIEditSkippedOperation,
  FolioAISignatureParty,
} from "./types";
import { diffWordSegments } from "./word-diff";

/**
 * The only editor surface the apply logic touches: a current `state`
 * and a `dispatch` that swaps in the next one. A live `EditorView`
 * satisfies this structurally, and so does a headless seam
 * (`{ state, dispatch: (tr) => { state = state.apply(tr); } }`) — the
 * apply path never reaches for anything DOM-bound on the view, which is
 * what lets the same operation applier drive both the React editor and
 * the server-side reviewer in `./headless`.
 */
export type FolioAIEditView = {
  state: EditorState;
  dispatch: (transaction: Transaction) => void;
};

type ApplyFolioAIEditOperationsOptions = {
  view: FolioAIEditView;
  snapshot: FolioAIEditSnapshot;
  operations: FolioAIEditOperation[];
  mode?: FolioAIEditApplyMode;
  author?: string;
  createCommentId?: (text: string) => number;
};

type ApplyFolioAIEditOperationsInternalOptions = ApplyFolioAIEditOperationsOptions & {
  revisionIdSeed?: number;
};

type ResolvedOperation = {
  operation: FolioAIEditOperation;
  from: number;
  to: number;
  blockFrom: number;
  blockTo: number;
  blockNode: PMNode;
  insertText?: string;
  tableRowInsertion?: TableRowInsertion;
  tableRowDeletion?: TableRowDeletion;
  tableColumnInsertion?: TableColumnInsertion;
  tableColumnDeletion?: TableColumnDeletion;
  tableCellMerge?: TableCellMerge;
  tableCellSplit?: TableCellSplit;
  commentId?: number;
  /**
   * Position in the input `operations` array, used as a secondary
   * sort key so same-position operations preserve the AI's logical
   * ordering when applied bottom-up.
   */
  originalIndex: number;
};

/**
 * Block attrs that identify a block instance (Word's
 * `w14:paraId` / `w14:textId`, future tracked-change identifiers).
 * Reuse during inheritFormatting would create duplicate IDs, so
 * we strip them when synthesising a sibling.
 */
const IDENTITY_BLOCK_ATTRS = new Set(["paraId", "textId"]);

const stripIdentityAttrs = (attrs: Record<string, unknown>) => {
  const next: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(attrs)) {
    if (IDENTITY_BLOCK_ATTRS.has(key)) {
      next[key] = null;
      continue;
    }
    next[key] = value;
  }
  return next;
};

const applyReplaceBlockStyleId = ({
  item,
  tr,
}: {
  item: ResolvedOperation;
  tr: Transaction;
}): Transaction => {
  if (item.operation.type !== "replaceBlock" || item.operation.styleId === undefined) {
    return tr;
  }

  const block = tr.doc.nodeAt(item.blockFrom);
  if (!block) {
    return tr;
  }

  return tr.setNodeMarkup(item.blockFrom, undefined, {
    ...block.attrs,
    styleId: item.operation.styleId,
  });
};

type ApplyInlineFormattingOptions = {
  tr: Transaction;
  schema: Schema;
  from: number;
  to: number;
  formatting: Extract<FolioAIEditOperation, { type: "formatRange" }>["formatting"];
};

const applyInlineFormatting = ({
  tr,
  schema,
  from,
  to,
  formatting,
}: ApplyInlineFormattingOptions): Transaction => {
  for (const [name, enabled] of Object.entries(formatting)) {
    const markType = schema.marks[name];
    if (!markType) {
      continue;
    }
    if (enabled) {
      tr.addMark(
        from,
        to,
        name === "underline" ? markType.create({ style: "single" }) : markType.create(),
      );
      continue;
    }
    tr.removeMark(from, to, markType);
  }
  return tr;
};

const formattingWouldChange = (
  marks: readonly Mark[],
  formatting: Extract<FolioAIEditOperation, { type: "formatRange" }>["formatting"],
): boolean =>
  Object.entries(formatting).some(([name, enabled]) => {
    const mark = marks.find((candidate) => candidate.type.name === name);
    if (!enabled) {
      return mark !== undefined;
    }
    if (name === "underline") {
      return mark?.attrs["style"] !== "single";
    }
    return mark === undefined;
  });

type ApplyTrackedInlineFormattingOptions = ApplyInlineFormattingOptions & {
  doc: PMNode;
  revisionId: number;
  author: string;
  date: string;
};

const applyTrackedInlineFormatting = ({
  tr,
  schema,
  doc,
  from,
  to,
  formatting,
  revisionId,
  author,
  date,
}: ApplyTrackedInlineFormattingOptions): Transaction => {
  const propertyChangeType = schema.marks["runPropertyChange"];
  if (!propertyChangeType) {
    return tr;
  }

  const segments: { from: number; to: number; changes: RunPropertyChange[] }[] = [];
  doc.nodesBetween(from, to, (node, pos) => {
    if (!node.isText || !formattingWouldChange(node.marks, formatting)) {
      return;
    }
    const segmentFrom = Math.max(from, pos);
    const segmentTo = Math.min(to, pos + node.nodeSize);
    const previousFormatting = marksToTextFormatting(node.marks);
    const existingMark = node.marks.find((mark) => mark.type === propertyChangeType);
    const existingChanges = existingMark
      ? expectRunPropertyChangeMarkAttrs(existingMark).changes
      : [];
    const change: RunPropertyChange = {
      type: "runPropertyChange",
      info: { id: revisionId, author, date },
      ...(Object.keys(previousFormatting).length > 0 ? { previousFormatting } : {}),
    };
    segments.push({
      from: segmentFrom,
      to: segmentTo,
      changes: [...existingChanges, change],
    });
  });

  if (segments.length === 0) {
    return tr;
  }
  applyInlineFormatting({ tr, schema, from, to, formatting });
  for (const segment of segments) {
    tr.addMark(segment.from, segment.to, propertyChangeType.create({ changes: segment.changes }));
  }
  return tr;
};

type LiveBlockEntry = { from: number; to: number; node: PMNode };

/**
 * Module-scoped monotonic counter for tracked-change revision ids.
 * Seeded once from `Date.now()` so ids are roughly time-ordered for
 * humans reading raw DOCX, then incremented per allocation. A bare
 * `Date.now()` seed per applyAIEditOperations call would collide
 * across batches that fire within the same millisecond (the panel's
 * Accept-all loop does exactly that — multiple calls in tight
 * succession). Reserving a contiguous range up front guarantees
 * uniqueness across overlapping calls in the same JS realm.
 */
let revisionIdCursor = Date.now() * 1000;
const nextRevisionSeed = (count: number): number => {
  // Each replace allocates two ids per op; insert/delete one each.
  // Reserve `count * 4` to be safely above any conceivable per-op
  // allocation (current max is 2). Returning the start of the
  // reserved range as the seed is enough — the caller bumps it.
  const start = revisionIdCursor;
  revisionIdCursor += Math.max(count, 1) * 4;
  return start;
};

/**
 * Walk the live doc once and bucket every textblock by its
 * normalised text hash. Resolution then maps each snapshot anchor
 * to the live block at the same ordinal among same-hash siblings —
 * unrelated edits that shift absolute positions no longer break
 * the lookup, and a sibling sharing text content with the target
 * doesn't trigger a false "changed" skip either.
 */
const collectLiveBlocksByHash = (doc: PMNode) => {
  const byHash = new Map<string, LiveBlockEntry[]>();
  doc.descendants((node, pos) => {
    if (!node.isTextblock) {
      return;
    }
    // Hash from the post-tracked-changes view so the snapshot
    // (taken with the same view) and live doc bucket the same
    // block under the same key. Otherwise a block mid-edit gets a
    // different hash than the snapshot recorded and the resolver
    // skips it as "changed".
    const cleanText = buildCleanBlockText(node, pos).text;
    const hash = hashFolioAIBlockText(normalizeFolioAIBlockText(cleanText));
    const bucket = byHash.get(hash) ?? [];
    bucket.push({ from: pos, to: pos + node.nodeSize, node });
    byHash.set(hash, bucket);
  });
  return byHash;
};

/**
 * Index live textblocks by their `w14:paraId`. Used by the resolver
 * to prefer a paraId-anchored lookup when the snapshot id encodes
 * one. Direct lookup avoids
 * the hash+ordinal failure mode where an earlier-in-document
 * duplicate of the same text gets picked instead of the actual
 * referenced paragraph.
 */
const collectLiveBlocksByParaId = (doc: PMNode) => {
  const byParaId = new Map<string, LiveBlockEntry>();
  doc.descendants((node, pos) => {
    if (!node.isTextblock) {
      return;
    }
    const paraId: unknown = node.attrs["paraId"];
    if (typeof paraId !== "string" || paraId.length === 0) {
      return;
    }
    // First-write-wins so an Enter-split duplicate (briefly co-existing
    // before the allocator re-issues) doesn't override the original.
    if (!byParaId.has(paraId)) {
      byParaId.set(paraId, { from: pos, to: pos + node.nodeSize, node });
    }
  });
  return byParaId;
};

/**
 * Insert ops (insertAfterBlock / insertBeforeBlock) target a block
 * by id. When that block lives inside a `tableCell`, naively
 * inserting at the block's `before` / `after` position drops the
 * new node INSIDE the cell — which is never what the model wants
 * when it says "insert after this paragraph". Escape outward to
 * the nearest enclosing `table` so the synthesized sibling lands
 * adjacent to the table as a doc-level peer.
 *
 * Returns the table's outer boundary positions when the block lies
 * inside one, otherwise `null` to keep the default block-adjacent
 * behaviour for paragraphs at the document root.
 */
type TableBoundary = { before: number; after: number };

const findEnclosingTableBoundary = (doc: PMNode, blockFrom: number): TableBoundary | null => {
  const resolved = doc.resolve(blockFrom);
  // Walk from the deepest ancestor outwards. `depth` is the
  // resolved-position's parent depth; the doc node sits at depth 0,
  // so the loop stops before trying to take a `before` of the doc.
  for (let depth = resolved.depth; depth > 0; depth--) {
    if (resolved.node(depth).type.name === "table") {
      return { before: resolved.before(depth), after: resolved.after(depth) };
    }
  }
  return null;
};

type TableRowInsertion = {
  tableStart: number;
  rowPosition: number;
  rowType: PMNode["type"];
  cells: readonly PMNode[];
  rowspanUpdates: readonly number[];
};

type TableRowTarget = {
  table: PMNode;
  tableStart: number;
  tablePosition: number;
  rowIndex: number;
  rowPosition: number;
};

type TableRowDeletion = Omit<TableRowTarget, "table">;

type TableColumnInsertion = {
  tablePosition: number;
  columnIndex: number;
};

type TableColumnDeletion = TableColumnInsertion;

type TableRectangle = {
  left: number;
  top: number;
  right: number;
  bottom: number;
};

type TableCellMerge = {
  tablePosition: number;
  rectangle: TableRectangle;
};

type TableCellSplit = TableCellMerge;

const getTableColumnCoordinateKey = ({
  tablePosition,
  columnIndex,
}: TableColumnInsertion): string => `${tablePosition}:${columnIndex}`;

type TableCellTarget = {
  tablePosition: number;
  cellPosition: number;
  cellEndPosition: number;
  leftColumnIndex: number;
  rightColumnIndex: number;
  topRowIndex: number;
  bottomRowIndex: number;
};

const findEnclosingTableRow = (doc: PMNode, blockFrom: number): TableRowTarget | null => {
  const resolved = doc.resolve(blockFrom);
  for (let rowDepth = resolved.depth; rowDepth > 0; rowDepth--) {
    if (resolved.node(rowDepth).type.spec["tableRole"] !== "row") {
      continue;
    }
    const tableDepth = rowDepth - 1;
    const table = resolved.node(tableDepth);
    if (table.type.spec["tableRole"] !== "table") {
      return null;
    }
    return {
      table,
      tableStart: resolved.start(tableDepth),
      tablePosition: resolved.before(tableDepth),
      rowIndex: resolved.index(tableDepth),
      rowPosition: resolved.before(rowDepth),
    };
  }
  return null;
};

const findEnclosingTableCell = (doc: PMNode, blockFrom: number): TableCellTarget | null => {
  const resolved = doc.resolve(blockFrom);
  for (let cellDepth = resolved.depth; cellDepth > 0; cellDepth--) {
    const cell = resolved.node(cellDepth);
    const tableRole = cell.type.spec["tableRole"];
    if (tableRole !== "cell" && tableRole !== "header_cell") {
      continue;
    }
    const rowDepth = cellDepth - 1;
    const tableDepth = rowDepth - 1;
    if (tableDepth < 0) {
      return null;
    }
    if (
      resolved.node(rowDepth).type.spec["tableRole"] !== "row" ||
      resolved.node(tableDepth).type.spec["tableRole"] !== "table"
    ) {
      return null;
    }
    const table = resolved.node(tableDepth);
    const tableStart = resolved.start(tableDepth);
    const cellPosition = resolved.before(cellDepth);
    const cellRect = TableMap.get(table).findCell(cellPosition - tableStart);
    return {
      tablePosition: resolved.before(tableDepth),
      cellPosition,
      cellEndPosition: cellPosition + cell.nodeSize,
      leftColumnIndex: cellRect.left,
      rightColumnIndex: cellRect.right,
      topRowIndex: cellRect.top,
      bottomRowIndex: cellRect.bottom,
    };
  }
  return null;
};

const tableRectanglesOverlap = (left: TableRectangle, right: TableRectangle): boolean =>
  left.left < right.right &&
  right.left < left.right &&
  left.top < right.bottom &&
  right.top < left.bottom;

const tableRectanglesEqual = (left: TableRectangle, right: TableRectangle): boolean =>
  left.left === right.left &&
  left.top === right.top &&
  left.right === right.right &&
  left.bottom === right.bottom;

const tableRectangleCutsMergedCell = (map: TableMap, rectangle: TableRectangle): boolean => {
  for (let row = rectangle.top; row < rectangle.bottom; row++) {
    const leftIndex = row * map.width + rectangle.left;
    const rightIndex = row * map.width + rectangle.right - 1;
    if (
      (rectangle.left > 0 && map.map[leftIndex] === map.map[leftIndex - 1]) ||
      (rectangle.right < map.width && map.map[rightIndex] === map.map[rightIndex + 1])
    ) {
      return true;
    }
  }
  for (let column = rectangle.left; column < rectangle.right; column++) {
    const topIndex = rectangle.top * map.width + column;
    const bottomIndex = (rectangle.bottom - 1) * map.width + column;
    if (
      (rectangle.top > 0 && map.map[topIndex] === map.map[topIndex - map.width]) ||
      (rectangle.bottom < map.height && map.map[bottomIndex] === map.map[bottomIndex + map.width])
    ) {
      return true;
    }
  }
  return false;
};

const isEmptyTableCell = (cell: PMNode): boolean =>
  cell.childCount === 1 &&
  cell.firstChild?.isTextblock === true &&
  cell.firstChild.childCount === 0;

const mergeTableRectangle = (
  tr: Transaction,
  tablePosition: number,
  table: PMNode,
  rectangle: TableRectangle,
): Transaction | null => {
  const map = TableMap.get(table);
  if (
    rectangle.left < 0 ||
    rectangle.top < 0 ||
    rectangle.right > map.width ||
    rectangle.bottom > map.height ||
    rectangle.left >= rectangle.right ||
    rectangle.top >= rectangle.bottom ||
    (rectangle.right - rectangle.left === 1 && rectangle.bottom - rectangle.top === 1) ||
    tableRectangleCutsMergedCell(map, rectangle)
  ) {
    return null;
  }

  const tableStart = tablePosition + 1;
  const seen = new Set<number>();
  const cells: { position: number; cell: PMNode }[] = [];
  let appendedContent = Fragment.empty;
  for (let row = rectangle.top; row < rectangle.bottom; row++) {
    for (let column = rectangle.left; column < rectangle.right; column++) {
      const cellPosition = map.map[row * map.width + column];
      if (cellPosition === undefined || seen.has(cellPosition)) {
        continue;
      }
      const cell = table.nodeAt(cellPosition);
      if (!cell) {
        return null;
      }
      seen.add(cellPosition);
      cells.push({ position: cellPosition, cell });
      if (cells.length > 1 && !isEmptyTableCell(cell)) {
        appendedContent = appendedContent.append(cell.content);
      }
    }
  }
  const merged = cells.at(0);
  if (!merged || cells.length < 2) {
    return null;
  }

  const colspan: unknown = merged.cell.attrs["colspan"];
  const rowspan: unknown = merged.cell.attrs["rowspan"];
  const colwidth: unknown = merged.cell.attrs["colwidth"];
  if (
    typeof colspan !== "number" ||
    !Number.isInteger(colspan) ||
    colspan < 1 ||
    typeof rowspan !== "number" ||
    !Number.isInteger(rowspan) ||
    rowspan < 1 ||
    (colwidth !== null &&
      colwidth !== undefined &&
      (!Array.isArray(colwidth) ||
        colwidth.length !== colspan ||
        !colwidth.every((width) => typeof width === "number")))
  ) {
    return null;
  }
  const mergedColspan = rectangle.right - rectangle.left;
  const nextColwidth = Array.isArray(colwidth) ? [...colwidth] : null;
  while (nextColwidth && nextColwidth.length < mergedColspan) {
    nextColwidth.push(0);
  }
  const mapFrom = tr.mapping.maps.length;
  for (const { position: cellPosition, cell } of cells.slice(1)) {
    const position = tr.mapping.slice(mapFrom).map(tableStart + cellPosition);
    tr = tr.delete(position, position + cell.nodeSize);
  }
  const absoluteMergedPosition = tableStart + merged.position;
  tr = tr.setNodeMarkup(absoluteMergedPosition, undefined, {
    ...merged.cell.attrs,
    colspan: mergedColspan,
    rowspan: rectangle.bottom - rectangle.top,
    colwidth: nextColwidth,
  });
  if (appendedContent.size > 0) {
    const contentEnd = absoluteMergedPosition + 1 + merged.cell.content.size;
    const contentStart = isEmptyTableCell(merged.cell) ? absoluteMergedPosition + 1 : contentEnd;
    tr = tr.replaceWith(contentStart, contentEnd, appendedContent);
  }
  return tr;
};

type MergeTrackedVerticalTableCellsOptions = {
  tr: Transaction;
  tablePosition: number;
  table: PMNode;
  rectangle: TableRectangle;
  revisionId: number;
  author: string;
  date: string;
};

const mergeTrackedVerticalTableCells = ({
  tr,
  tablePosition,
  table,
  rectangle,
  revisionId,
  author,
  date,
}: MergeTrackedVerticalTableCellsOptions): Transaction | null => {
  const map = TableMap.get(table);
  if (
    rectangle.right - rectangle.left !== 1 ||
    rectangle.bottom - rectangle.top < 2 ||
    rectangle.left < 0 ||
    rectangle.top < 0 ||
    rectangle.right > map.width ||
    rectangle.bottom > map.height ||
    tableRectangleCutsMergedCell(map, rectangle)
  ) {
    return null;
  }

  const cells: { position: number; cell: PMNode }[] = [];
  for (let row = rectangle.top; row < rectangle.bottom; row++) {
    const position = map.map[row * map.width + rectangle.left];
    if (position === undefined) {
      return null;
    }
    const cell = table.nodeAt(position);
    if (
      !cell ||
      !tableRectanglesEqual(map.findCell(position), {
        left: rectangle.left,
        top: row,
        right: rectangle.right,
        bottom: row + 1,
      })
    ) {
      return null;
    }
    const attrs = expectTableCellAttrs(cell);
    if (
      attrs.colspan !== 1 ||
      attrs.rowspan !== 1 ||
      attrs.cellMarker !== undefined ||
      attrs._docxVMergeContinuationCells !== undefined ||
      attrs._preserveVMergeRestart === true ||
      attrs._originalFormatting?.vMerge !== undefined
    ) {
      return null;
    }
    cells.push({ position, cell });
  }

  const origin = cells.at(0);
  if (
    !origin ||
    cells.length !== rectangle.bottom - rectangle.top ||
    cells.some(({ cell }) => cell.type !== origin.cell.type) ||
    cells.slice(1).some(({ cell }) => !isEmptyTableCell(cell))
  ) {
    return null;
  }

  const continuationCells = cells.slice(1).map(({ cell }) => {
    const continuation = standaloneTableCellFromProseMirror(cell);
    return {
      ...continuation,
      formatting: {
        ...continuation.formatting,
        vMerge: "continue" as const,
      },
      structuralChange: {
        type: "tableCellMerge" as const,
        info: { id: revisionId, author, date },
        verticalMerge: "continue" as const,
        verticalMergeOriginal: "rest" as const,
      },
    };
  });

  const nextTr = mergeTableRectangle(tr, tablePosition, table, rectangle);
  if (!nextTr) {
    return null;
  }
  nextTr.setNodeAttribute(
    tablePosition + 1 + origin.position,
    "_docxVMergeContinuationCells",
    continuationCells,
  );
  markStructuralChange(nextTr);
  return nextTr;
};

const splitTableRectangle = (
  tr: Transaction,
  tablePosition: number,
  table: PMNode,
  rectangle: TableRectangle,
): Transaction | null => {
  const map = TableMap.get(table);
  if (
    rectangle.left < 0 ||
    rectangle.top < 0 ||
    rectangle.right > map.width ||
    rectangle.bottom > map.height ||
    rectangle.left >= rectangle.right ||
    rectangle.top >= rectangle.bottom ||
    (rectangle.right - rectangle.left === 1 && rectangle.bottom - rectangle.top === 1)
  ) {
    return null;
  }
  const cellPosition = map.map[rectangle.top * map.width + rectangle.left];
  if (cellPosition === undefined) {
    return null;
  }
  const cellRectangle = map.findCell(cellPosition);
  if (!tableRectanglesEqual(cellRectangle, rectangle)) {
    return null;
  }
  const cell = table.nodeAt(cellPosition);
  if (!cell) {
    return null;
  }

  const colspan: unknown = cell.attrs["colspan"];
  const rowspan: unknown = cell.attrs["rowspan"];
  const colwidth: unknown = cell.attrs["colwidth"];
  if (
    typeof colspan !== "number" ||
    !Number.isInteger(colspan) ||
    colspan !== rectangle.right - rectangle.left ||
    typeof rowspan !== "number" ||
    !Number.isInteger(rowspan) ||
    rowspan !== rectangle.bottom - rectangle.top ||
    (colwidth !== null &&
      colwidth !== undefined &&
      (!Array.isArray(colwidth) ||
        colwidth.length !== colspan ||
        !colwidth.every((width) => typeof width === "number")))
  ) {
    return null;
  }

  const baseAttrs = {
    ...cell.attrs,
    colspan: 1,
    rowspan: 1,
    _preserveVMergeRestart: null,
    _docxVMergeContinuationCells: null,
  };
  const attrsByColumn = Array.from({ length: colspan }, (_, index) => ({
    ...baseAttrs,
    colwidth:
      Array.isArray(colwidth) && typeof colwidth[index] === "number" && colwidth[index] > 0
        ? [colwidth[index]]
        : null,
  }));
  const tableStart = tablePosition + 1;
  const insertions: { position: number; cell: PMNode }[] = [];
  for (let row = rectangle.top; row < rectangle.bottom; row++) {
    let position = map.positionAt(row, rectangle.left, table);
    if (row === rectangle.top) {
      position += cell.nodeSize;
    }
    for (let column = rectangle.left; column < rectangle.right; column++) {
      if (row === rectangle.top && column === rectangle.left) {
        continue;
      }
      const attrs = attrsByColumn[column - rectangle.left];
      const nextCell = attrs ? cell.type.createAndFill(attrs) : null;
      if (!nextCell) {
        return null;
      }
      insertions.push({ position: tableStart + position, cell: nextCell });
    }
  }

  const mapFrom = tr.mapping.maps.length;
  for (const insertion of insertions) {
    tr = tr.insert(tr.mapping.slice(mapFrom).map(insertion.position, 1), insertion.cell);
  }
  return tr.setNodeMarkup(tableStart + cellPosition, undefined, attrsByColumn.at(0));
};

type SplitTrackedVerticalTableCellOptions = {
  tr: Transaction;
  tablePosition: number;
  table: PMNode;
  rectangle: TableRectangle;
  revisionId: number;
  author: string;
  date: string;
};

const splitTrackedVerticalTableCell = ({
  tr,
  tablePosition,
  table,
  rectangle,
  revisionId,
  author,
  date,
}: SplitTrackedVerticalTableCellOptions): Transaction | null => {
  const map = TableMap.get(table);
  if (
    rectangle.right - rectangle.left !== 1 ||
    rectangle.bottom - rectangle.top < 2 ||
    rectangle.left < 0 ||
    rectangle.top < 0 ||
    rectangle.right > map.width ||
    rectangle.bottom > map.height
  ) {
    return null;
  }
  const relativeCellPosition = map.map[rectangle.top * map.width + rectangle.left];
  if (relativeCellPosition === undefined) {
    return null;
  }
  const cellRectangle = map.findCell(relativeCellPosition);
  if (!tableRectanglesEqual(cellRectangle, rectangle)) {
    return null;
  }
  const cell = table.nodeAt(relativeCellPosition);
  if (!cell) {
    return null;
  }
  const attrs = expectTableCellAttrs(cell);
  if (
    attrs.colspan !== 1 ||
    attrs.rowspan !== rectangle.bottom - rectangle.top ||
    attrs.cellMarker !== undefined
  ) {
    return null;
  }
  const continuationCells = attrs._docxVMergeContinuationCells;
  if (
    continuationCells !== undefined &&
    continuationCells.length !== rectangle.bottom - rectangle.top - 1
  ) {
    return null;
  }

  const marker = {
    kind: "merge" as const,
    info: { revisionId, author, date },
    verticalMergeOriginal: "continue" as const,
  };
  const insertedCells: PMNode[] = [];
  for (let index = 0; index < rectangle.bottom - rectangle.top - 1; index++) {
    const source = continuationCells?.[index];
    if (source?.structuralChange !== undefined) {
      return null;
    }
    const insertedCell = source
      ? trackedSplitCellFromStoredSource(cell, source, marker)
      : cell.type.createAndFill({
          ...cell.attrs,
          rowspan: 1,
          cellMarker: marker,
          _originalFormatting: formattingWithoutVerticalMerge(attrs._originalFormatting),
          _preserveVMergeRestart: null,
          _docxVMergeContinuationCells: null,
        });
    if (!insertedCell) {
      return null;
    }
    insertedCells.push(insertedCell);
  }

  const tableStart = tablePosition + 1;
  const mapFrom = tr.mapping.maps.length;
  for (const [index, insertedCell] of insertedCells.entries()) {
    const row = rectangle.top + index + 1;
    const position = tableStart + map.positionAt(row, rectangle.left, table);
    tr = tr.insert(tr.mapping.slice(mapFrom).map(position, 1), insertedCell);
  }
  tr = tr.setNodeMarkup(tableStart + relativeCellPosition, undefined, {
    ...cell.attrs,
    rowspan: 1,
    _originalFormatting: formattingWithoutVerticalMerge(attrs._originalFormatting),
    _preserveVMergeRestart: null,
    _docxVMergeContinuationCells: null,
  });
  markStructuralChange(tr);
  return tr;
};

type TrackedSplitCellMarker = {
  kind: "merge";
  info: {
    revisionId: number;
    author: string;
    date: string;
  };
  verticalMergeOriginal: "continue";
};

const trackedSplitCellFromStoredSource = (
  origin: PMNode,
  source: TableCell,
  marker: TrackedSplitCellMarker,
): PMNode | null => {
  const formatting = formattingWithoutVerticalMerge(source.formatting);
  const restoredSource: TableCell = {
    type: "tableCell",
    ...(formatting ? { formatting } : {}),
    ...(source.propertyChanges ? { propertyChanges: source.propertyChanges } : {}),
    content: source.content,
  };
  const converted = standaloneTableCellToProseMirror(
    restoredSource,
    origin.type.name === "tableHeader" ? "tableHeader" : "tableCell",
  );
  const compatible = Result.try(() => origin.type.schema.nodeFromJSON(converted.toJSON())).unwrapOr(
    null,
  );
  if (!compatible) {
    return null;
  }
  return compatible.type.create(
    {
      ...compatible.attrs,
      rowspan: 1,
      cellMarker: marker,
      _preserveVMergeRestart: null,
      _docxVMergeContinuationCells: null,
    },
    compatible.content,
  );
};

const formattingWithoutVerticalMerge = (
  formatting: TableCellFormatting | undefined,
): TableCellFormatting | undefined => {
  if (!formatting) {
    return undefined;
  }
  const next = { ...formatting };
  delete next.vMerge;
  return next;
};

const findTableColumnInsertion = (
  doc: PMNode,
  blockFrom: number,
  position: "after" | "before",
): (TableColumnInsertion & { boundaryPosition: number }) | null => {
  const target = findEnclosingTableCell(doc, blockFrom);
  if (!target) {
    return null;
  }
  return {
    tablePosition: target.tablePosition,
    columnIndex: position === "before" ? target.leftColumnIndex : target.rightColumnIndex,
    boundaryPosition: position === "before" ? target.cellPosition : target.cellEndPosition,
  };
};

const deleteTableNode = (
  tr: Transaction,
  tablePosition: number,
  table: PMNode,
): Transaction | null => {
  const tableEnd = tablePosition + table.nodeSize;
  const resolved = tr.doc.resolve(tablePosition);
  const parent = resolved.parent;
  const tableIndex = resolved.index();
  if (parent.canReplace(tableIndex, tableIndex + 1)) {
    return tr.delete(tablePosition, tableEnd);
  }
  const emptyParagraph = tr.doc.type.schema.nodes["paragraph"]?.createAndFill();
  if (
    !emptyParagraph ||
    !parent.canReplaceWith(tableIndex, tableIndex + 1, emptyParagraph.type, emptyParagraph.marks)
  ) {
    return null;
  }
  return tr.replaceWith(tablePosition, tableEnd, emptyParagraph);
};

const buildTableRowInsertion = (
  map: TableMap,
  table: PMNode,
  tableStart: number,
  rowIndex: number,
): TableRowInsertion | null => {
  const cells: PMNode[] = [];
  const rowspanUpdates: number[] = [];
  let referenceRow: number | null = rowIndex > 0 ? -1 : 0;
  if (rowIsHeader(map, table, rowIndex + referenceRow)) {
    referenceRow = rowIndex === 0 || rowIndex === map.height ? null : 0;
  }
  for (let col = 0, index = map.width * rowIndex; col < map.width; col++, index++) {
    if (rowIndex > 0 && rowIndex < map.height && map.map[index] === map.map[index - map.width]) {
      const cellPosition = map.map[index];
      if (cellPosition === undefined) {
        return null;
      }
      const cell = table.nodeAt(cellPosition);
      if (!cell) {
        return null;
      }
      const colspan: unknown = cell.attrs["colspan"];
      if (typeof colspan !== "number" || colspan < 1) {
        return null;
      }
      rowspanUpdates.push(cellPosition);
      col += colspan - 1;
      index += colspan - 1;
      continue;
    }
    const referencePosition =
      referenceRow === null ? undefined : map.map[index + referenceRow * map.width];
    const cellType =
      referencePosition === undefined
        ? tableNodeTypes(table.type.schema).cell
        : table.nodeAt(referencePosition)?.type;
    const cell = cellType?.createAndFill();
    if (!cell) {
      return null;
    }
    cells.push(cell);
  }
  if (cells.length === 0) {
    return null;
  }
  let rowPosition = tableStart;
  for (let index = 0; index < rowIndex; index++) {
    rowPosition += table.child(index).nodeSize;
  }
  return {
    tableStart,
    rowPosition,
    rowType: tableNodeTypes(table.type.schema).row,
    cells,
    rowspanUpdates,
  };
};

const findTableRowInsertion = (
  doc: PMNode,
  blockFrom: number,
  position: "after" | "before",
): TableRowInsertion | null => {
  const target = findEnclosingTableRow(doc, blockFrom);
  if (!target) {
    return null;
  }
  const map = TableMap.get(target.table);
  const rowIndex = position === "before" ? target.rowIndex : target.rowIndex + 1;
  return buildTableRowInsertion(map, target.table, target.tableStart, rowIndex);
};

const populateTableRow = (row: PMNode, cellTexts: readonly string[] | undefined): PMNode => {
  if (cellTexts === undefined) {
    return row;
  }
  const cells: PMNode[] = [];
  row.forEach((cell, _offset, index) => {
    const paragraph = cell.firstChild;
    if (!paragraph?.isTextblock) {
      cells.push(cell);
      return;
    }
    const text = cellTexts[index] ?? "";
    const content = text.length > 0 ? row.type.schema.text(text) : null;
    const nextParagraph = paragraph.type.create(stripIdentityAttrs(paragraph.attrs), content);
    cells.push(cell.type.create(cell.attrs, nextParagraph));
  });
  return row.type.create(row.attrs, cells);
};

const getInsertedPhysicalColumnRows = (map: TableMap, columnIndex: number): number[] => {
  const rows: number[] = [];
  for (let row = 0; row < map.height; row++) {
    const index = row * map.width + columnIndex;
    const crossesColspan =
      columnIndex > 0 && columnIndex < map.width && map.map[index - 1] === map.map[index];
    if (!crossesColspan) {
      rows.push(row);
    }
  }
  return rows;
};

type TableColumnInsertionAction =
  | { type: "insertCell"; position: number; cell: PMNode }
  | { type: "setColspan"; position: number; attrs: Record<string, unknown> };

type BuildTableColumnInsertionActionsOptions = {
  map: TableMap;
  table: PMNode;
  columnIndex: number;
  cellTexts: readonly string[] | undefined;
};

const expandTableCellColspan = (
  cell: PMNode,
  colspanOffset: number,
): Record<string, unknown> | null => {
  const colspan: unknown = cell.attrs["colspan"];
  const rowspan: unknown = cell.attrs["rowspan"];
  const colwidth: unknown = cell.attrs["colwidth"];
  if (
    typeof colspan !== "number" ||
    !Number.isInteger(colspan) ||
    colspan < 1 ||
    typeof rowspan !== "number" ||
    !Number.isInteger(rowspan) ||
    rowspan < 1 ||
    !Number.isInteger(colspanOffset) ||
    colspanOffset < 0 ||
    colspanOffset > colspan
  ) {
    return null;
  }
  if (colwidth !== null && !Array.isArray(colwidth)) {
    return null;
  }
  if (Array.isArray(colwidth) && !colwidth.every((width) => typeof width === "number")) {
    return null;
  }
  const nextColwidth = Array.isArray(colwidth) ? [...colwidth] : null;
  nextColwidth?.splice(colspanOffset, 0, 0);
  return {
    ...cell.attrs,
    colspan: colspan + 1,
    colwidth: nextColwidth,
  };
};

const populateTableCell = (cell: PMNode, text: string): PMNode | null => {
  const paragraph = cell.firstChild;
  if (!paragraph?.isTextblock) {
    return null;
  }
  const content = text.length > 0 ? cell.type.schema.text(text) : null;
  const nextParagraph = paragraph.type.create(stripIdentityAttrs(paragraph.attrs), content);
  return cell.type.create(cell.attrs, nextParagraph);
};

const buildTableColumnInsertionActions = ({
  map,
  table,
  columnIndex,
  cellTexts,
}: BuildTableColumnInsertionActionsOptions): TableColumnInsertionAction[] | null => {
  if (columnIndex < 0 || columnIndex > map.width) {
    return null;
  }
  const physicalRows = getInsertedPhysicalColumnRows(map, columnIndex);
  if ((cellTexts?.length ?? 0) > physicalRows.length) {
    return null;
  }
  let referenceColumn: number | null = columnIndex > 0 ? -1 : 0;
  if (columnIsHeader(map, table, columnIndex + referenceColumn)) {
    referenceColumn = columnIndex === 0 || columnIndex === map.width ? null : 0;
  }
  const actions: TableColumnInsertionAction[] = [];
  let physicalCellIndex = 0;
  for (let row = 0; row < map.height; row++) {
    const index = row * map.width + columnIndex;
    if (columnIndex > 0 && columnIndex < map.width && map.map[index - 1] === map.map[index]) {
      const position = map.map[index];
      if (position === undefined) {
        return null;
      }
      const cell = table.nodeAt(position);
      if (!cell) {
        return null;
      }
      const colspanOffset = columnIndex - map.colCount(position);
      const attrs = expandTableCellColspan(cell, colspanOffset);
      if (!attrs) {
        return null;
      }
      actions.push({
        type: "setColspan",
        position,
        attrs,
      });
      row += Number(cell.attrs["rowspan"]) - 1;
      continue;
    }
    const referencePosition =
      referenceColumn === null ? undefined : map.map[index + referenceColumn];
    const cellType =
      referencePosition === undefined
        ? tableNodeTypes(table.type.schema).cell
        : table.nodeAt(referencePosition)?.type;
    const cell = cellType?.createAndFill();
    if (!cell) {
      return null;
    }
    const populatedCell = populateTableCell(cell, cellTexts?.[physicalCellIndex] ?? "");
    if (!populatedCell) {
      return null;
    }
    actions.push({
      type: "insertCell",
      position: map.positionAt(row, columnIndex, table),
      cell: populatedCell,
    });
    physicalCellIndex++;
  }
  return actions;
};

type GetTableColumnDeletionCellPositionsOptions = {
  map: TableMap;
  table: PMNode;
  columnIndex: number;
};

const getTableColumnDeletionCellPositions = ({
  map,
  table,
  columnIndex,
}: GetTableColumnDeletionCellPositionsOptions): number[] | null => {
  if (columnIndex < 0 || columnIndex >= map.width) {
    return null;
  }
  const positions = new Set<number>();
  for (let row = 0; row < map.height; row++) {
    const position = map.map[row * map.width + columnIndex];
    if (position === undefined) {
      return null;
    }
    const cell = table.nodeAt(position);
    const tableRole = cell?.type.spec["tableRole"];
    if (
      !cell ||
      (tableRole !== "cell" && tableRole !== "header_cell") ||
      cell.attrs["colspan"] !== 1 ||
      cell.attrs["cellMarker"] != null
    ) {
      return null;
    }
    positions.add(position);
  }
  return [...positions];
};

/**
 * The snapshot recorded an `hashOccurrenceCount` per anchor but
 * not which ordinal within that bucket the block was — recompute
 * on demand from the snapshot's anchor map. Stable iteration
 * (object insertion order) means anchors with the same hash come
 * out in document order, which is what we want.
 */
const ordinalAmongSameHash = (snapshot: FolioAIEditSnapshot, blockId: string): number => {
  const target = snapshot.anchors[blockId];
  if (!target) {
    return -1;
  }
  let ordinal = 0;
  for (const anchor of Object.values(snapshot.anchors)) {
    if (anchor.id === blockId) {
      return ordinal;
    }
    if (anchor.textHash === target.textHash) {
      ordinal += 1;
    }
  }
  return -1;
};

/**
 * Length of the underscore signature rule. Mirrors the constant
 * in `docx-core/legal-source/compile.ts` so the on-screen signature
 * line matches what `create-document` produces.
 */
const SIGNATURE_LINE = "_".repeat(28);

type BuildSignatureTableNodeOptions = {
  schema: Schema;
  parties: readonly FolioAISignatureParty[];
};

/**
 * Build a borderless PM table mirroring docx-core's
 * `signatureTable` helper: one row, one cell per party, each cell
 * containing party name (bold), two spacer paragraphs, a signature
 * rule, then optional signatory and italic title lines.
 *
 * Returns `null` when the editor schema is missing one of the
 * required node types — callers should surface the op as a skip
 * rather than crashing.
 */
const buildSignatureTableNode = ({
  schema,
  parties,
}: BuildSignatureTableNodeOptions): PMNode | null => {
  const paragraphType = schema.nodes["paragraph"];
  const cellType = schema.nodes["tableCell"];
  const rowType = schema.nodes["tableRow"];
  const tableType = schema.nodes["table"];
  if (!paragraphType || !cellType || !rowType || !tableType) {
    return null;
  }
  const boldType = schema.marks["bold"];
  const italicType = schema.marks["italic"];

  const buildParagraph = (
    text: string,
    options: {
      styleId: string;
      bold?: boolean;
      italic?: boolean;
    },
  ): PMNode => {
    const marks: Mark[] = [];
    if (options.bold && boldType) {
      marks.push(boldType.create());
    }
    if (options.italic && italicType) {
      marks.push(italicType.create());
    }
    const content = text.length > 0 ? schema.text(text, marks) : null;
    return paragraphType.create({ styleId: options.styleId }, content);
  };

  const buildCell = (party: FolioAISignatureParty): PMNode => {
    const cellContent: PMNode[] = [
      buildParagraph(party.name, { styleId: "SignatureParty", bold: true }),
      buildParagraph("", { styleId: "SignatureSpacer" }),
      buildParagraph("", { styleId: "SignatureSpacer" }),
      buildParagraph(SIGNATURE_LINE, { styleId: "SignatureRule" }),
    ];
    if (party.signatory && party.signatory.length > 0) {
      cellContent.push(buildParagraph(party.signatory, { styleId: "SignatureField" }));
    }
    if (party.title && party.title.length > 0) {
      cellContent.push(
        buildParagraph(party.title, {
          styleId: "SignatureField",
          italic: true,
        }),
      );
    }
    // Cell attrs left to schema defaults — column widths and
    // borders are decided by the table renderer; we just need the
    // structural shape.
    return cellType.create({}, cellContent);
  };

  const cells = parties.map(buildCell);
  const row = rowType.create({}, cells);
  return tableType.create({}, row);
};

/**
 * Build the inline content for an inserted or replaced block, promoting the
 * model's `**bold**` / `***bold italic***` markdown into real Word marks (the
 * edit-tool schema has no inline-format channel, so the model improvises with
 * markdown that would otherwise land as literal asterisks). Falls back to a
 * single verbatim text node when no emphasis is present, so plain prose is
 * never reshaped. `baseMarks` (insertion / comment) ride on every run.
 */
const buildEmphasisInlineContent = (
  schema: Schema,
  text: string,
  baseMarks: readonly Mark[],
): PMNode[] => {
  const runs = parseInlineEmphasisRuns(text);
  if (!runs.some((run) => run.bold || run.italic)) {
    return [schema.text(text, [...baseMarks])];
  }
  const boldType = schema.marks["bold"];
  const italicType = schema.marks["italic"];
  const nodes: PMNode[] = [];
  for (const run of runs) {
    if (run.text.length === 0) {
      continue;
    }
    const marks: Mark[] = [...baseMarks];
    if (run.bold && boldType) {
      marks.push(boldType.create());
    }
    if (run.italic && italicType) {
      marks.push(italicType.create());
    }
    nodes.push(schema.text(run.text, marks));
  }
  return nodes.length > 0 ? nodes : [schema.text(text, [...baseMarks])];
};

const applyFolioAIEditOperationsInternal = ({
  view,
  snapshot,
  operations,
  mode = "tracked-changes",
  author = "AI",
  createCommentId,
  revisionIdSeed,
}: ApplyFolioAIEditOperationsInternalOptions): FolioAIEditApplyResult => {
  const applied: FolioAIEditAppliedOperation[] = [];
  const skipped: FolioAIEditSkippedOperation[] = [];
  const resolved: ResolvedOperation[] = [];
  const insertionType = view.state.schema.marks["insertion"];
  const deletionType = view.state.schema.marks["deletion"];
  const commentType = view.state.schema.marks["comment"];
  const claimedTableRows = new Set<string>();
  const claimedTableColumns = new Set<string>();

  if (mode === "tracked-changes" && (!insertionType || !deletionType)) {
    return {
      applied,
      skipped: operations.map((operation) => ({
        id: operation.id,
        reason: "unsupportedBlock",
      })),
    };
  }

  // Build the live-block indexes once per batch so individual op
  // resolutions don't each re-walk the doc. ParaId-anchored ids are
  // resolved against `liveBlocksByParaId`; the hash bucket is only
  // the fallback for ordinal-encoded snapshot ids.
  const liveBlocks = collectLiveBlocksByHash(view.state.doc);
  const liveBlocksByParaId = collectLiveBlocksByParaId(view.state.doc);

  for (const [index, operation] of operations.entries()) {
    const commentText = getOperationCommentText(operation);
    if (commentText !== undefined && (!commentType || createCommentId === undefined)) {
      skipped.push({ id: operation.id, reason: "unsupportedBlock" });
      continue;
    }

    const resolution = resolveOperation({
      snapshot,
      operation,
      liveBlocks,
      liveBlocksByParaId,
      doc: view.state.doc,
    });
    if (resolution.type === "skip") {
      skipped.push({ id: operation.id, reason: resolution.reason });
      continue;
    }

    const deletion = resolution.operation.tableRowDeletion;
    if (deletion) {
      const rowKey = `${deletion.tablePosition}:${deletion.rowIndex}`;
      if (claimedTableRows.has(rowKey)) {
        skipped.push({ id: operation.id, reason: "noopOperation" });
        continue;
      }
      claimedTableRows.add(rowKey);
    }

    const columnDeletion = resolution.operation.tableColumnDeletion;
    if (columnDeletion) {
      const columnKey = getTableColumnCoordinateKey(columnDeletion);
      if (claimedTableColumns.has(columnKey)) {
        skipped.push({ id: operation.id, reason: "noopOperation" });
        continue;
      }
      claimedTableColumns.add(columnKey);
    }

    const commentId = commentText !== undefined ? createCommentId?.(commentText) : undefined;
    resolved.push({
      ...resolution.operation,
      originalIndex: index,
      ...(commentId !== undefined && { commentId }),
    });
  }

  const tableStructureMutations = new Set<number>();
  for (const item of resolved) {
    const tablePosition =
      item.tableRowInsertion?.tableStart !== undefined
        ? item.tableRowInsertion.tableStart - 1
        : (item.tableRowDeletion?.tablePosition ??
          item.tableColumnInsertion?.tablePosition ??
          item.tableColumnDeletion?.tablePosition);
    if (tablePosition !== undefined) {
      tableStructureMutations.add(tablePosition);
    }
  }

  const mergeTables = new Set(
    resolved.flatMap((item) => (item.tableCellMerge ? [item.tableCellMerge.tablePosition] : [])),
  );
  const splitTables = new Set(
    resolved.flatMap((item) => (item.tableCellSplit ? [item.tableCellSplit.tablePosition] : [])),
  );
  const mergeRectanglesByTable = new Map<number, TableRectangle[]>();
  const splitRectanglesByTable = new Map<number, TableRectangle[]>();
  const executableResolved: ResolvedOperation[] = [];
  for (const item of resolved) {
    const merge = item.tableCellMerge;
    if (merge) {
      if (
        tableStructureMutations.has(merge.tablePosition) ||
        splitTables.has(merge.tablePosition)
      ) {
        skipped.push({ id: item.operation.id, reason: "unsupportedBlock" });
        continue;
      }
      const claimedRectangles = mergeRectanglesByTable.get(merge.tablePosition) ?? [];
      const duplicate = claimedRectangles.some((rectangle) =>
        tableRectanglesEqual(rectangle, merge.rectangle),
      );
      if (duplicate) {
        skipped.push({ id: item.operation.id, reason: "noopOperation" });
        continue;
      }
      const overlap = claimedRectangles.some((rectangle) =>
        tableRectanglesOverlap(rectangle, merge.rectangle),
      );
      if (overlap) {
        skipped.push({ id: item.operation.id, reason: "unsupportedBlock" });
        continue;
      }
      claimedRectangles.push(merge.rectangle);
      mergeRectanglesByTable.set(merge.tablePosition, claimedRectangles);
      executableResolved.push(item);
      continue;
    }
    const split = item.tableCellSplit;
    if (split) {
      if (
        tableStructureMutations.has(split.tablePosition) ||
        mergeTables.has(split.tablePosition)
      ) {
        skipped.push({ id: item.operation.id, reason: "unsupportedBlock" });
        continue;
      }
      const claimedRectangles = splitRectanglesByTable.get(split.tablePosition) ?? [];
      if (claimedRectangles.some((rectangle) => tableRectanglesEqual(rectangle, split.rectangle))) {
        skipped.push({ id: item.operation.id, reason: "noopOperation" });
        continue;
      }
      if (
        claimedRectangles.some((rectangle) => tableRectanglesOverlap(rectangle, split.rectangle))
      ) {
        skipped.push({ id: item.operation.id, reason: "unsupportedBlock" });
        continue;
      }
      claimedRectangles.push(split.rectangle);
      splitRectanglesByTable.set(split.tablePosition, claimedRectangles);
    }
    executableResolved.push(item);
  }

  if (executableResolved.length === 0) {
    return { applied, skipped };
  }

  let tr = view.state.tr;
  let revisionSeed = revisionIdSeed ?? nextRevisionSeed(executableResolved.length);
  const date = new Date().toISOString();
  const insertedColumnCounts = new Map<string, number>();

  // Sort right-to-left so each tr.insert / tr.delete leaves earlier
  // positions intact. Column insertions run before deletions at the
  // same snapshot coordinate; the deletion path accounts for those
  // inserted columns so it still removes the original target. Other
  // ties use reverse input order so repeated insertions retain their
  // requested sequence.
  for (const item of executableResolved.toSorted((left, right) => {
    const leftCellShape = left.tableCellMerge ?? left.tableCellSplit;
    const rightCellShape = right.tableCellMerge ?? right.tableCellSplit;
    if (!leftCellShape && rightCellShape) {
      return -1;
    }
    if (leftCellShape && !rightCellShape) {
      return 1;
    }
    if (leftCellShape && rightCellShape) {
      if (leftCellShape.tablePosition !== rightCellShape.tablePosition) {
        return rightCellShape.tablePosition - leftCellShape.tablePosition;
      }
      if (leftCellShape.rectangle.bottom !== rightCellShape.rectangle.bottom) {
        return rightCellShape.rectangle.bottom - leftCellShape.rectangle.bottom;
      }
      if (leftCellShape.rectangle.right !== rightCellShape.rectangle.right) {
        return rightCellShape.rectangle.right - leftCellShape.rectangle.right;
      }
      return right.originalIndex - left.originalIndex;
    }
    const leftColumn = left.tableColumnInsertion ?? left.tableColumnDeletion;
    const rightColumn = right.tableColumnInsertion ?? right.tableColumnDeletion;
    if (!leftColumn && rightColumn) {
      return -1;
    }
    if (leftColumn && !rightColumn) {
      return 1;
    }
    if (leftColumn && rightColumn) {
      if (leftColumn.tablePosition !== rightColumn.tablePosition) {
        return rightColumn.tablePosition - leftColumn.tablePosition;
      }
      if (leftColumn.columnIndex !== rightColumn.columnIndex) {
        return rightColumn.columnIndex - leftColumn.columnIndex;
      }
      const leftIsInsertion = left.tableColumnInsertion !== undefined;
      const rightIsInsertion = right.tableColumnInsertion !== undefined;
      if (leftIsInsertion !== rightIsInsertion) {
        return leftIsInsertion ? -1 : 1;
      }
      return right.originalIndex - left.originalIndex;
    }
    if (left.from !== right.from) {
      return right.from - left.from;
    }
    return right.originalIndex - left.originalIndex;
  })) {
    const commentMark =
      item.commentId !== undefined && commentType
        ? commentType.create({ commentId: item.commentId })
        : null;

    // Snapshot the transaction's step count so we can detect when an
    // operation produced zero document changes and report it as a
    // skipped no-op instead of a phantom "applied" entry. This caught
    // the silent accept-failure bug where a replaceInBlock on a
    // block with pending tracked changes computed the wrong source
    // text and the diff produced no marks; the panel said "accepted"
    // but the doc was untouched.
    const stepsBefore = tr.steps.length;
    let appliedRevisionIds: number[] | undefined;

    switch (item.operation.type) {
      case "replaceInBlock":
      case "replaceRange": {
        const revisionIdDelete = revisionSeed++;
        const revisionIdInsert = revisionSeed++;
        tr = applyTextReplacement({
          tr,
          item,
          mode,
          author,
          date,
          revisionIdDelete,
          revisionIdInsert,
          commentMark,
        });
        if (mode === "tracked-changes") {
          appliedRevisionIds = [revisionIdDelete, revisionIdInsert];
        }
        break;
      }
      case "commentOnRange": {
        if (commentMark) {
          tr = tr.addMark(item.from, item.to, commentMark);
        }
        break;
      }
      case "formatRange": {
        if (mode === "tracked-changes") {
          const revisionId = revisionSeed++;
          tr = applyTrackedInlineFormatting({
            tr,
            schema: view.state.schema,
            doc: tr.doc,
            from: item.from,
            to: item.to,
            formatting: item.operation.formatting,
            revisionId,
            author,
            date,
          });
          appliedRevisionIds = [revisionId];
          break;
        }
        tr = applyInlineFormatting({
          tr,
          schema: view.state.schema,
          from: item.from,
          to: item.to,
          formatting: item.operation.formatting,
        });
        break;
      }
      case "replaceBlock": {
        const revisionIdDelete = revisionSeed++;
        const revisionIdInsert = revisionSeed++;
        // Default to preserving formatting (existing behaviour);
        // when explicitly disabled and we're in direct mode, swap
        // the whole block node for a fresh paragraph that drops
        // all block-level attrs. tracked-changes mode keeps the
        // attrs because the visible diff is text-only.
        if (item.operation.preserveFormatting === false && mode === "direct") {
          const replacement = item.operation.text;
          const paragraphType = view.state.schema.nodes["paragraph"];
          if (paragraphType) {
            const replaceAttrs: Record<string, unknown> | null =
              item.operation.styleId !== undefined ? { styleId: item.operation.styleId } : null;
            const node = paragraphType.create(
              replaceAttrs,
              replacement.length === 0
                ? null
                : buildEmphasisInlineContent(view.state.schema, replacement, []),
            );
            tr = tr.replaceWith(item.blockFrom, item.blockTo, node);
            break;
          }
        }
        // Direct mode, formatting preserved (the default): when the replacement
        // carries inline emphasis, rebuild the block node keeping its own attrs
        // so `**bold**` becomes real marks. The tracked-changes path can't carry
        // inline marks through its word-diff redline, so it still strips them;
        // plain replacements fall through to the text-only swap below unchanged.
        if (mode === "direct" && hasInlineEmphasis(item.operation.text)) {
          const attrs: Record<string, unknown> =
            item.operation.styleId !== undefined
              ? { ...item.blockNode.attrs, styleId: item.operation.styleId }
              : { ...item.blockNode.attrs };
          const node = item.blockNode.type.create(
            attrs,
            buildEmphasisInlineContent(
              view.state.schema,
              item.operation.text,
              commentMark ? [commentMark] : [],
            ),
          );
          tr = tr.replaceWith(item.blockFrom, item.blockTo, node);
          break;
        }
        tr = applyTextReplacement({
          tr,
          item,
          mode,
          author,
          date,
          revisionIdDelete,
          revisionIdInsert,
          commentMark,
        });
        tr = applyReplaceBlockStyleId({ item, tr });
        if (mode === "tracked-changes") {
          appliedRevisionIds = [revisionIdDelete, revisionIdInsert];
        }
        break;
      }
      case "insertAfterBlock":
      case "insertBeforeBlock": {
        if (
          mode === "tracked-changes" &&
          item.operation.pageBreakBefore === true &&
          (!item.insertText || item.insertText.length === 0)
        ) {
          skipped.push({
            id: item.operation.id,
            reason: "unsupportedMode",
          });
          continue;
        }

        const marks = [];
        if (mode === "tracked-changes" && insertionType) {
          const revisionId = revisionSeed++;
          marks.push(
            insertionType.create({
              revisionId,
              author,
              date,
            }),
          );
          appliedRevisionIds = [revisionId];
        }
        if (commentMark) {
          marks.push(commentMark);
        }
        const content =
          item.insertText && item.insertText.length > 0
            ? buildEmphasisInlineContent(view.state.schema, item.insertText, marks)
            : null;
        // Inherit formatting attrs (listMarker, styleId, …) from
        // the source block but never reuse identity attrs — a new
        // paragraph must get fresh paraId/textId so trackers don't
        // collide.
        const baseAttrs =
          item.operation.inheritFormatting === false
            ? {}
            : stripIdentityAttrs(item.blockNode.attrs);
        const attrs: Record<string, unknown> = { ...baseAttrs };
        if (item.operation.pageBreakBefore === true) {
          attrs["pageBreakBefore"] = true;
        }
        if (item.operation.styleId !== undefined) {
          // Heading / clause style ids (e.g. ClauseHeading1) take
          // precedence over the source block's style so the
          // inserted paragraph renders as the requested kind, not
          // as a clone of the anchor.
          attrs["styleId"] = item.operation.styleId;
          // A heading is logically a fresh block; drop list marker
          // attrs that would otherwise leak from the anchor and
          // render the heading as a list item.
          if (item.operation.inheritFormatting !== false) {
            attrs["listMarker"] = null;
            attrs["listMarkerHidden"] = null;
            attrs["listLevelNumFmts"] = null;
            attrs["listLevelStarts"] = null;
            attrs["listAbstractNumId"] = null;
            attrs["listStartOverride"] = null;
          }
        }
        const node = item.blockNode.type.create(attrs, content);
        tr = tr.insert(item.from, node);
        break;
      }
      case "insertSignatureTable": {
        if (mode === "tracked-changes") {
          skipped.push({
            id: item.operation.id,
            reason: "unsupportedMode",
          });
          continue;
        }

        const node = buildSignatureTableNode({
          schema: view.state.schema,
          parties: item.operation.parties,
        });
        if (!node) {
          skipped.push({
            id: item.operation.id,
            reason: "unsupportedBlock",
          });
          continue;
        }
        // Tables don't carry tracked-change marks here — the
        // table structure itself is the insert, and inline
        // insertion marks on the paragraph runs inside would
        // double-up with the structural addition. Apply directly;
        // tracked-changes for new tables is a separate future
        // concern.
        tr = tr.insert(item.from, node);
        break;
      }
      case "insertTableRow": {
        const insertion = item.tableRowInsertion;
        if (!insertion) {
          skipped.push({
            id: item.operation.id,
            reason: "unsupportedBlock",
          });
          continue;
        }
        if (mode === "tracked-changes" && insertion.rowspanUpdates.length > 0) {
          skipped.push({
            id: item.operation.id,
            reason: "unsupportedBlock",
          });
          continue;
        }
        const liveRowspanUpdates: { position: number; cell: PMNode; rowspan: number }[] = [];
        let invalidRowspanUpdate = false;
        for (const updatePosition of insertion.rowspanUpdates) {
          const cellPosition = insertion.tableStart + updatePosition;
          const liveCell = tr.doc.nodeAt(cellPosition);
          if (!liveCell) {
            invalidRowspanUpdate = true;
            break;
          }
          const rowspan: unknown = liveCell.attrs["rowspan"];
          if (typeof rowspan !== "number") {
            invalidRowspanUpdate = true;
            break;
          }
          liveRowspanUpdates.push({ position: cellPosition, cell: liveCell, rowspan });
        }
        if (invalidRowspanUpdate) {
          skipped.push({
            id: item.operation.id,
            reason: "unsupportedBlock",
          });
          continue;
        }
        for (const update of liveRowspanUpdates) {
          tr = tr.setNodeMarkup(update.position, undefined, {
            ...update.cell.attrs,
            rowspan: update.rowspan + 1,
          });
        }
        let rowAttrs: Record<string, unknown> | null = null;
        if (mode === "tracked-changes") {
          const revisionId = revisionSeed++;
          rowAttrs = {
            trIns: {
              revisionId,
              author,
              date,
            },
          };
          appliedRevisionIds = [revisionId];
        }
        const row = insertion.rowType.create(rowAttrs, insertion.cells);
        tr = tr.insert(insertion.rowPosition, populateTableRow(row, item.operation.cellTexts));
        break;
      }
      case "insertTableColumn": {
        const insertion = item.tableColumnInsertion;
        const tablePosition = insertion ? tr.mapping.map(insertion.tablePosition, 1) : null;
        const table = tablePosition === null ? null : tr.doc.nodeAt(tablePosition);
        if (
          !insertion ||
          tablePosition === null ||
          !table ||
          table.type.spec["tableRole"] !== "table"
        ) {
          skipped.push({
            id: item.operation.id,
            reason: "unsupportedBlock",
          });
          continue;
        }
        const map = TableMap.get(table);
        const actions = buildTableColumnInsertionActions({
          map,
          table,
          columnIndex: insertion.columnIndex,
          cellTexts: item.operation.cellTexts,
        });
        if (
          !actions ||
          (mode === "tracked-changes" && actions.some(({ type }) => type === "setColspan"))
        ) {
          skipped.push({
            id: item.operation.id,
            reason: "unsupportedBlock",
          });
          continue;
        }
        const revisionId = mode === "tracked-changes" ? revisionSeed++ : null;
        if (revisionId !== null) {
          appliedRevisionIds = [revisionId];
        }
        const mapFrom = tr.mapping.maps.length;
        for (const action of actions) {
          const position = tr.mapping.slice(mapFrom).map(tablePosition + 1 + action.position);
          if (action.type === "insertCell") {
            const cell =
              revisionId === null
                ? action.cell
                : action.cell.type.create(
                    {
                      ...action.cell.attrs,
                      cellMarker: {
                        kind: "ins",
                        info: {
                          revisionId,
                          author,
                          date,
                        },
                      },
                    },
                    action.cell.content,
                  );
            tr = tr.insert(position, cell);
            continue;
          }
          tr = tr.setNodeMarkup(position, undefined, action.attrs);
        }
        const columnKey = getTableColumnCoordinateKey(insertion);
        insertedColumnCounts.set(columnKey, (insertedColumnCounts.get(columnKey) ?? 0) + 1);
        break;
      }
      case "deleteTableColumn": {
        const deletion = item.tableColumnDeletion;
        const tablePosition = deletion ? tr.mapping.map(deletion.tablePosition, 1) : null;
        const table = tablePosition === null ? null : tr.doc.nodeAt(tablePosition);
        if (
          !deletion ||
          tablePosition === null ||
          !table ||
          table.type.spec["tableRole"] !== "table"
        ) {
          skipped.push({
            id: item.operation.id,
            reason: "unsupportedBlock",
          });
          continue;
        }
        const map = TableMap.get(table);
        const columnKey = getTableColumnCoordinateKey(deletion);
        const columnIndex = deletion.columnIndex + (insertedColumnCounts.get(columnKey) ?? 0);
        if (columnIndex < 0 || columnIndex >= map.width) {
          skipped.push({
            id: item.operation.id,
            reason: "unsupportedBlock",
          });
          continue;
        }
        if (mode === "tracked-changes") {
          const cellPositions = getTableColumnDeletionCellPositions({
            map,
            table,
            columnIndex,
          });
          if (!cellPositions) {
            skipped.push({
              id: item.operation.id,
              reason: "unsupportedBlock",
            });
            continue;
          }
          const revisionId = revisionSeed++;
          for (const cellPosition of cellPositions) {
            tr = tr.setNodeAttribute(tablePosition + 1 + cellPosition, "cellMarker", {
              kind: "del",
              info: {
                revisionId,
                author,
                date,
              },
            });
          }
          markStructuralChange(tr);
          appliedRevisionIds = [revisionId];
          break;
        }
        if (map.width === 1) {
          const nextTr = deleteTableNode(tr, tablePosition, table);
          if (!nextTr) {
            skipped.push({
              id: item.operation.id,
              reason: "unsupportedBlock",
            });
            continue;
          }
          tr = nextTr;
          break;
        }
        removeColumn(
          tr,
          {
            map,
            table,
            tableStart: tablePosition + 1,
            left: columnIndex,
            top: 0,
            right: columnIndex + 1,
            bottom: map.height,
          },
          columnIndex,
        );
        break;
      }
      case "mergeTableCells": {
        const merge = item.tableCellMerge;
        const tablePosition = merge ? tr.mapping.map(merge.tablePosition, 1) : null;
        const table = tablePosition === null ? null : tr.doc.nodeAt(tablePosition);
        if (
          !merge ||
          tablePosition === null ||
          !table ||
          table.type.spec["tableRole"] !== "table"
        ) {
          skipped.push({
            id: item.operation.id,
            reason: "unsupportedBlock",
          });
          continue;
        }
        const revisionId = mode === "tracked-changes" ? revisionSeed : null;
        const nextTr =
          revisionId === null
            ? mergeTableRectangle(tr, tablePosition, table, merge.rectangle)
            : mergeTrackedVerticalTableCells({
                tr,
                tablePosition,
                table,
                rectangle: merge.rectangle,
                revisionId,
                author,
                date,
              });
        if (!nextTr) {
          skipped.push({
            id: item.operation.id,
            reason: "unsupportedBlock",
          });
          continue;
        }
        tr = nextTr;
        if (revisionId !== null) {
          revisionSeed++;
          appliedRevisionIds = [revisionId];
        }
        break;
      }
      case "splitTableCell": {
        const split = item.tableCellSplit;
        const tablePosition = split ? tr.mapping.map(split.tablePosition, 1) : null;
        const table = tablePosition === null ? null : tr.doc.nodeAt(tablePosition);
        if (
          !split ||
          tablePosition === null ||
          !table ||
          table.type.spec["tableRole"] !== "table"
        ) {
          skipped.push({
            id: item.operation.id,
            reason: "unsupportedBlock",
          });
          continue;
        }
        const revisionId = mode === "tracked-changes" ? revisionSeed : null;
        const nextTr =
          revisionId === null
            ? splitTableRectangle(tr, tablePosition, table, split.rectangle)
            : splitTrackedVerticalTableCell({
                tr,
                tablePosition,
                table,
                rectangle: split.rectangle,
                revisionId,
                author,
                date,
              });
        if (!nextTr) {
          skipped.push({
            id: item.operation.id,
            reason: "unsupportedBlock",
          });
          continue;
        }
        tr = nextTr;
        if (revisionId !== null) {
          revisionSeed++;
          appliedRevisionIds = [revisionId];
        }
        break;
      }
      case "deleteTableRow": {
        const deletion = item.tableRowDeletion;
        const rowPosition = deletion ? tr.mapping.map(deletion.rowPosition) : null;
        const row = rowPosition === null ? null : tr.doc.nodeAt(rowPosition);
        if (mode === "tracked-changes") {
          if (
            !deletion ||
            rowPosition === null ||
            !row ||
            row.type.spec["tableRole"] !== "row" ||
            row.attrs["trIns"] != null ||
            row.attrs["trDel"] != null
          ) {
            skipped.push({
              id: item.operation.id,
              reason: "unsupportedBlock",
            });
            continue;
          }
          const revisionId = revisionSeed++;
          tr = tr.setNodeAttribute(rowPosition, "trDel", {
            revisionId,
            author,
            date,
          });
          markStructuralChange(tr);
          appliedRevisionIds = [revisionId];
          break;
        }
        const table = deletion ? tr.doc.nodeAt(deletion.tablePosition) : null;
        if (
          !deletion ||
          !table ||
          table.type.spec["tableRole"] !== "table" ||
          deletion.rowIndex >= table.childCount
        ) {
          skipped.push({
            id: item.operation.id,
            reason: "unsupportedBlock",
          });
          continue;
        }
        if (table.childCount === 1) {
          const nextTr = deleteTableNode(tr, deletion.tablePosition, table);
          if (!nextTr) {
            skipped.push({
              id: item.operation.id,
              reason: "unsupportedBlock",
            });
            continue;
          }
          tr = nextTr;
          break;
        }
        const map = TableMap.get(table);
        removeRow(
          tr,
          {
            map,
            table,
            tableStart: deletion.tableStart,
            left: 0,
            top: deletion.rowIndex,
            right: map.width,
            bottom: deletion.rowIndex + 1,
          },
          deletion.rowIndex,
        );
        break;
      }
      case "deleteBlock": {
        if (mode === "direct") {
          tr = tr.delete(item.blockFrom, item.blockTo);
          break;
        }

        if (deletionType) {
          const revisionId = revisionSeed++;
          tr = tr.addMark(
            item.from,
            item.to,
            deletionType.create({
              revisionId,
              author,
              date,
            }),
          );
          appliedRevisionIds = [revisionId];
        }
        if (commentMark) {
          tr = tr.addMark(item.from, item.to, commentMark);
        }
        break;
      }
      case "commentOnBlock": {
        if (commentMark) {
          tr = tr.addMark(item.from, item.to, commentMark);
        }
        break;
      }
      default:
        break;
    }

    // commentOnBlock is intentionally a doc-mutation-free op
    // (adds a mark; that DOES count as a step) but covers the
    // edge case where the comment mark is missing. Treat any op
    // that emitted zero transaction steps as a no-op skip.
    if (tr.steps.length === stepsBefore) {
      skipped.push({ id: item.operation.id, reason: "noopOperation" });
      continue;
    }

    // Surface the primary id (first one) on the legacy `revisionId`
    // field so callers that just need a stable scroll/visual
    // reference keep working. The full set is on `revisionIds` for
    // accept/reject paths that must clear every mark belonging to
    // this op.
    applied.push({
      id: item.operation.id,
      ...(item.commentId !== undefined && { commentId: item.commentId }),
      ...(appliedRevisionIds !== undefined &&
        appliedRevisionIds[0] !== undefined && {
          revisionId: appliedRevisionIds[0],
          revisionIds: appliedRevisionIds,
        }),
    });
  }

  if (tr.docChanged) {
    view.dispatch(tr);
  }

  return { applied, skipped };
};

export const applyFolioAIEditOperations = (
  options: ApplyFolioAIEditOperationsOptions,
): FolioAIEditApplyResult => applyFolioAIEditOperationsInternal(options);

export const previewFolioAIEditOperations = (
  options: ApplyFolioAIEditOperationsOptions,
): FolioAIEditApplyResult => {
  const { view, createCommentId, ...applyOptions } = options;
  let previewCommentId = -1;
  const previewView: FolioAIEditView = {
    state: view.state,
    dispatch: (transaction) => {
      previewView.state = previewView.state.apply(transaction);
    },
  };
  const result = applyFolioAIEditOperationsInternal({
    ...applyOptions,
    view: previewView,
    ...(createCommentId !== undefined && {
      createCommentId: () => previewCommentId--,
    }),
    revisionIdSeed: -1_000_000_000,
  });
  return {
    applied: result.applied.map(({ id }) => ({ id })),
    skipped: result.skipped,
  };
};

type TextReplacementOptions = {
  tr: Transaction;
  item: ResolvedOperation;
  mode: FolioAIEditApplyMode;
  author: string;
  date: string;
  /**
   * Distinct revision id used for the deletion-side marks.
   * fromProseDoc treats a single revisionId carrying BOTH ins and
   * del marks as a Word "moveTo/moveFrom" pair on serialization,
   * which is wrong for an AI replace — so the engine allocates one
   * id for the deletion side and a separate one for the insertion
   * side of the same operation.
   */
  revisionIdDelete: number;
  /** Distinct revision id used for the insertion-side marks. */
  revisionIdInsert: number;
  commentMark: Mark | null;
};

const applyTextReplacement = ({
  tr,
  item,
  mode,
  author,
  date,
  revisionIdDelete,
  revisionIdInsert,
  commentMark,
}: TextReplacementOptions): Transaction => {
  let nextTr = tr;
  const replacement = stripInlineEmphasisMarkers(
    (() => {
      if (item.operation.type === "replaceInBlock" || item.operation.type === "replaceRange") {
        return item.operation.replace;
      }
      if (item.operation.type === "replaceBlock") {
        return item.operation.text;
      }
      return "";
    })(),
  );

  if (mode === "direct") {
    // Partial-block replacement carrying inline emphasis: rebuild the matched
    // range as real bold/italic runs instead of stripping the markers to plain
    // text. Full-block replaceBlock in direct mode is intercepted earlier and
    // never reaches here with emphasis; the tracked-changes path below still
    // strips, since its word-diff redline can't carry inline marks.
    if (
      (item.operation.type === "replaceInBlock" || item.operation.type === "replaceRange") &&
      hasInlineEmphasis(item.operation.replace)
    ) {
      const content = buildEmphasisInlineContent(
        nextTr.doc.type.schema,
        item.operation.replace,
        commentMark ? [commentMark] : [],
      );
      return nextTr.replaceWith(item.from, item.to, content);
    }
    nextTr = nextTr.insertText(replacement, item.from, item.to);
    if (commentMark && replacement.length > 0) {
      nextTr = nextTr.addMark(item.from, item.from + replacement.length, commentMark);
    }
    return nextTr;
  }

  const insertionType = nextTr.doc.type.schema.marks["insertion"];
  const deletionType = nextTr.doc.type.schema.marks["deletion"];
  const delAttrs = { revisionId: revisionIdDelete, author, date };
  const insAttrs = { revisionId: revisionIdInsert, author, date };

  // Word-level diff is only safe when the source range maps to PM
  // positions losslessly. The block must have no atomic inline
  // nodes (hard breaks, inline images) — those break textContent /
  // PM-position alignment in ways the offsets array can't resolve.
  //
  // Existing tracked-change marks ARE handled here: we walk PM
  // positions through `cleanBlock.offsets[]` (built from the
  // post-tracked-changes view), so each clean-text char anchors at
  // the right live position even when the block has pending
  // deletion runs interleaved between surviving chars. Naively
  // accumulating `cursor += seg.text.length` would skip the gap
  // introduced by those deletion runs and write marks onto the
  // wrong live characters — the silent accept-failure bug.
  const blockHasOnlyTextChildren =
    item.blockNode.content.size === item.blockNode.textContent.length;
  const cleanBlock = blockHasOnlyTextChildren
    ? buildCleanBlockText(item.blockNode, item.blockFrom)
    : null;
  let sourceText: string | null = null;
  let sourceCleanStart = 0;
  if (cleanBlock !== null) {
    if (item.operation.type === "replaceInBlock") {
      sourceText = item.operation.find;
      sourceCleanStart = cleanBlock.text.indexOf(item.operation.find);
      if (sourceCleanStart === -1) {
        sourceText = null;
      }
    } else if (item.operation.type === "replaceRange") {
      sourceCleanStart = item.operation.range.startOffset;
      sourceText = cleanBlock.text.slice(sourceCleanStart, item.operation.range.endOffset);
    } else if (item.operation.type === "replaceBlock") {
      sourceText = cleanBlock.text;
      sourceCleanStart = 0;
    }
  }

  if (sourceText !== null && cleanBlock !== null) {
    const segments = diffWordSegments(sourceText, replacement);
    const offsets = cleanBlock.offsets;
    const offsetAt = (cleanOffset: number): number | null => offsets[cleanOffset] ?? null;

    type Step =
      | { kind: "del"; from: number; to: number }
      | { kind: "ins"; at: number; text: string };
    const steps: Step[] = [];
    // Cursor walks SOURCE-text offsets (within `sourceText`), then
    // we translate to PM positions through offsets[]. This survives
    // gaps caused by existing deletion-marked runs in the live doc.
    let cursor = 0;
    let allPositionsResolved = true;
    for (const seg of segments) {
      if (seg.type === "equal") {
        cursor += seg.text.length;
        continue;
      }
      if (seg.type === "del") {
        const pmFrom = offsetAt(sourceCleanStart + cursor);
        const pmTo = offsetAt(sourceCleanStart + cursor + seg.text.length);
        if (pmFrom === null || pmTo === null) {
          allPositionsResolved = false;
          break;
        }
        steps.push({ kind: "del", from: pmFrom, to: pmTo });
        cursor += seg.text.length;
        continue;
      }
      const pmAt = offsetAt(sourceCleanStart + cursor);
      if (pmAt === null) {
        allPositionsResolved = false;
        break;
      }
      steps.push({ kind: "ins", at: pmAt, text: seg.text });
    }

    if (allPositionsResolved) {
      // Apply right-to-left so earlier steps' source positions stay
      // valid after later steps mutate the doc.
      for (const step of steps.toReversed()) {
        if (step.kind === "del" && deletionType) {
          nextTr = nextTr.addMark(step.from, step.to, deletionType.create(delAttrs));
          if (commentMark) {
            nextTr = nextTr.addMark(step.from, step.to, commentMark);
          }
          continue;
        }
        if (step.kind === "ins" && insertionType) {
          nextTr = nextTr.insertText(step.text, step.at, step.at);
          nextTr = nextTr.addMark(
            step.at,
            step.at + step.text.length,
            insertionType.create(insAttrs),
          );
          if (commentMark) {
            nextTr = nextTr.addMark(step.at, step.at + step.text.length, commentMark);
          }
        }
      }
      return nextTr;
    }
    // else fall through to the coarse del+ins path below: the
    // offsets array didn't cover one of our boundaries, which only
    // happens for edge cases at the trailing block boundary.
  }

  if (replacement.length > 0 && insertionType) {
    nextTr = nextTr.insertText(replacement, item.to, item.to);
    nextTr = nextTr.addMark(item.to, item.to + replacement.length, insertionType.create(insAttrs));
    if (commentMark) {
      nextTr = nextTr.addMark(item.to, item.to + replacement.length, commentMark);
    }
  }

  if (item.to > item.from && deletionType) {
    nextTr = nextTr.addMark(item.from, item.to, deletionType.create(delAttrs));
    if (commentMark && replacement.length === 0) {
      nextTr = nextTr.addMark(item.from, item.to, commentMark);
    }
  }

  return nextTr;
};

type ResolvedBase = Omit<ResolvedOperation, "originalIndex" | "commentId">;

type ResolveOperationArgs = {
  snapshot: FolioAIEditSnapshot;
  operation: FolioAIEditOperation;
  liveBlocks: Map<string, LiveBlockEntry[]>;
  liveBlocksByParaId: Map<string, LiveBlockEntry>;
  doc: PMNode;
};

type ResolveStableBlockArgs = Omit<ResolveOperationArgs, "operation" | "doc"> & {
  blockId: string;
};

type StableBlockResolution = {
  type: "resolved";
  blockNode: PMNode;
  blockFrom: number;
  blockTo: number;
  cleanBlock: ReturnType<typeof buildCleanBlockText>;
  currentText: string;
  currentTextHash: string;
};

const resolveStableBlock = ({
  snapshot,
  blockId,
  liveBlocks,
  liveBlocksByParaId,
}: ResolveStableBlockArgs): StableBlockResolution | OperationResolutionSkip => {
  const anchor = snapshot.anchors[blockId];
  if (!anchor) {
    return { type: "skip", reason: "missingBlock" };
  }

  const encodedParaId = getFolioParaIdFromBlockId(blockId);
  let live: LiveBlockEntry | undefined;
  if (encodedParaId !== null) {
    live = liveBlocksByParaId.get(encodedParaId);
    if (!live) {
      return { type: "skip", reason: "missingBlock" };
    }
  } else {
    const ordinal = ordinalAmongSameHash(snapshot, blockId);
    if (ordinal < 0) {
      return { type: "skip", reason: "missingBlock" };
    }
    live = liveBlocks.get(anchor.textHash)?.[ordinal];
  }
  if (!live || !live.node.isTextblock) {
    return { type: "skip", reason: "changedBlock" };
  }

  const cleanBlock = buildCleanBlockText(live.node, live.from);
  const currentText = cleanBlock.text;
  const currentTextHash = hashFolioAIBlockText(normalizeFolioAIBlockText(currentText));
  if (currentTextHash !== anchor.textHash) {
    return { type: "skip", reason: "changedBlock" };
  }
  return {
    type: "resolved",
    blockNode: live.node,
    blockFrom: live.from,
    blockTo: live.to,
    cleanBlock,
    currentText,
    currentTextHash,
  };
};

const resolveOperation = ({
  snapshot,
  operation,
  liveBlocks,
  liveBlocksByParaId,
  doc,
}: ResolveOperationArgs):
  | { type: "resolved"; operation: ResolvedBase }
  | OperationResolutionSkip => {
  const blockId =
    operation.type === "replaceRange" ||
    operation.type === "commentOnRange" ||
    operation.type === "formatRange"
      ? operation.range.blockId
      : operation.blockId;
  const primaryBlock = resolveStableBlock({
    snapshot,
    blockId,
    liveBlocks,
    liveBlocksByParaId,
  });
  if (primaryBlock.type === "skip") {
    return primaryBlock;
  }
  const { blockNode, blockFrom, blockTo, cleanBlock, currentText, currentTextHash } = primaryBlock;
  if (
    operation.precondition !== undefined &&
    currentTextHash !== operation.precondition.blockTextHash
  ) {
    return { type: "skip", reason: "preconditionFailed" };
  }

  if (
    operation.type === "replaceRange" ||
    operation.type === "commentOnRange" ||
    operation.type === "formatRange"
  ) {
    const { startOffset, endOffset, selectedTextHash } = operation.range;
    const from = cleanBlock.offsets[startOffset];
    const to = cleanBlock.offsets[endOffset];
    if (from === undefined || to === undefined) {
      return { type: "skip", reason: "staleRange" };
    }
    const selectedText = currentText.slice(startOffset, endOffset);
    if (hashFolioAIBlockText(selectedText) !== selectedTextHash) {
      return { type: "skip", reason: "staleRange" };
    }
    if (operation.type === "replaceRange" && selectedText === operation.replace) {
      return { type: "skip", reason: "noopOperation" };
    }
    return {
      type: "resolved",
      operation: {
        operation,
        from,
        to,
        blockFrom,
        blockTo,
        blockNode,
      },
    };
  }

  if (operation.type === "insertAfterBlock" || operation.type === "insertBeforeBlock") {
    // An empty `text` is normally an error, but a page-break-only
    // paragraph is legitimate (the user just wants whitespace +
    // forced break). Letting it through with no inline content
    // means the inserted block renders as an empty paragraph that
    // starts a new page.
    if (operation.text.length === 0 && operation.pageBreakBefore !== true) {
      return { type: "skip", reason: "emptyOperation" };
    }
    // If the anchor lives inside a `tableCell`, the model meant
    // "place the new block adjacent to the table", not "stuff it
    // into the cell". Override the insertion bounds to the table's
    // outer boundary so the synthesized sibling lands as a peer of
    // the table at doc level.
    const tableBoundary = findEnclosingTableBoundary(doc, blockFrom);
    const isInsertAfter = operation.type === "insertAfterBlock";
    let insertFrom: number;
    if (tableBoundary) {
      insertFrom = isInsertAfter ? tableBoundary.after : tableBoundary.before;
    } else {
      insertFrom = isInsertAfter ? blockTo : blockFrom;
    }
    return {
      type: "resolved",
      operation: {
        operation,
        from: insertFrom,
        to: insertFrom,
        blockFrom,
        blockTo,
        blockNode,
        insertText: operation.text,
      },
    };
  }

  if (operation.type === "insertSignatureTable") {
    if (operation.parties.length === 0) {
      return { type: "skip", reason: "emptyOperation" };
    }
    const position = operation.position ?? "after";
    const tableBoundary = findEnclosingTableBoundary(doc, blockFrom);
    let insertFrom: number;
    if (tableBoundary) {
      insertFrom = position === "after" ? tableBoundary.after : tableBoundary.before;
    } else {
      insertFrom = position === "after" ? blockTo : blockFrom;
    }
    return {
      type: "resolved",
      operation: {
        operation,
        from: insertFrom,
        to: insertFrom,
        blockFrom,
        blockTo,
        blockNode,
      },
    };
  }

  if (operation.type === "insertTableRow") {
    const position = operation.position ?? "after";
    const insertion = findTableRowInsertion(doc, blockFrom, position);
    if (!insertion || (operation.cellTexts?.length ?? 0) > insertion.cells.length) {
      return { type: "skip", reason: "unsupportedBlock" };
    }
    return {
      type: "resolved",
      operation: {
        operation,
        from: insertion.rowPosition,
        to: insertion.rowPosition,
        blockFrom,
        blockTo,
        blockNode,
        tableRowInsertion: insertion,
      },
    };
  }

  if (operation.type === "deleteTableRow") {
    const target = findEnclosingTableRow(doc, blockFrom);
    if (!target) {
      return { type: "skip", reason: "unsupportedBlock" };
    }
    return {
      type: "resolved",
      operation: {
        operation,
        from: target.rowPosition,
        to: target.rowPosition,
        blockFrom,
        blockTo,
        blockNode,
        tableRowDeletion: {
          tableStart: target.tableStart,
          tablePosition: target.tablePosition,
          rowIndex: target.rowIndex,
          rowPosition: target.rowPosition,
        },
      },
    };
  }

  if (operation.type === "insertTableColumn") {
    const position = operation.position ?? "after";
    const insertion = findTableColumnInsertion(doc, blockFrom, position);
    if (!insertion) {
      return { type: "skip", reason: "unsupportedBlock" };
    }
    const { boundaryPosition, ...tableColumnInsertion } = insertion;
    return {
      type: "resolved",
      operation: {
        operation,
        from: boundaryPosition,
        to: boundaryPosition,
        blockFrom,
        blockTo,
        blockNode,
        tableColumnInsertion,
      },
    };
  }

  if (operation.type === "deleteTableColumn") {
    const target = findEnclosingTableCell(doc, blockFrom);
    if (!target) {
      return { type: "skip", reason: "unsupportedBlock" };
    }
    return {
      type: "resolved",
      operation: {
        operation,
        from: target.cellPosition,
        to: target.cellPosition,
        blockFrom,
        blockTo,
        blockNode,
        tableColumnDeletion: {
          tablePosition: target.tablePosition,
          columnIndex: target.leftColumnIndex,
        },
      },
    };
  }

  if (operation.type === "mergeTableCells") {
    const startCell = findEnclosingTableCell(doc, blockFrom);
    if (!startCell) {
      return { type: "skip", reason: "unsupportedBlock" };
    }
    const table = doc.nodeAt(startCell.tablePosition);
    if (!table || table.type.spec["tableRole"] !== "table") {
      return { type: "skip", reason: "unsupportedBlock" };
    }
    const map = TableMap.get(table);
    let rectangle: TableRectangle;
    let endCellPosition: number;
    let endCellEndPosition: number;
    if (operation.rowCount !== undefined) {
      rectangle = {
        left: startCell.leftColumnIndex,
        top: startCell.topRowIndex,
        right: startCell.rightColumnIndex,
        bottom: startCell.topRowIndex + operation.rowCount,
      };
      if (
        rectangle.right - rectangle.left !== 1 ||
        rectangle.bottom > map.height ||
        operation.rowCount < 2
      ) {
        return { type: "skip", reason: "unsupportedBlock" };
      }
      const endRelativePosition = map.map[(rectangle.bottom - 1) * map.width + rectangle.left];
      const endCell = endRelativePosition === undefined ? null : table.nodeAt(endRelativePosition);
      if (!endCell || endRelativePosition === undefined) {
        return { type: "skip", reason: "unsupportedBlock" };
      }
      endCellPosition = startCell.tablePosition + 1 + endRelativePosition;
      endCellEndPosition = endCellPosition + endCell.nodeSize;
    } else {
      const endTarget = resolveStableBlock({
        snapshot,
        blockId: operation.endBlockId,
        liveBlocks,
        liveBlocksByParaId,
      });
      if (endTarget.type === "skip") {
        return endTarget;
      }
      const endCell = findEnclosingTableCell(doc, endTarget.blockFrom);
      if (!endCell || startCell.tablePosition !== endCell.tablePosition) {
        return { type: "skip", reason: "unsupportedBlock" };
      }
      rectangle = {
        left: Math.min(startCell.leftColumnIndex, endCell.leftColumnIndex),
        top: Math.min(startCell.topRowIndex, endCell.topRowIndex),
        right: Math.max(startCell.rightColumnIndex, endCell.rightColumnIndex),
        bottom: Math.max(startCell.bottomRowIndex, endCell.bottomRowIndex),
      };
      endCellPosition = endCell.cellPosition;
      endCellEndPosition = endCell.cellEndPosition;
    }
    if (
      (rectangle.right - rectangle.left === 1 && rectangle.bottom - rectangle.top === 1) ||
      tableRectangleCutsMergedCell(map, rectangle)
    ) {
      return {
        type: "skip",
        reason:
          rectangle.right - rectangle.left === 1 && rectangle.bottom - rectangle.top === 1
            ? "noopOperation"
            : "unsupportedBlock",
      };
    }
    return {
      type: "resolved",
      operation: {
        operation,
        from: Math.min(startCell.cellPosition, endCellPosition),
        to: Math.max(startCell.cellEndPosition, endCellEndPosition),
        blockFrom,
        blockTo,
        blockNode,
        tableCellMerge: {
          tablePosition: startCell.tablePosition,
          rectangle,
        },
      },
    };
  }

  if (operation.type === "splitTableCell") {
    const cell = findEnclosingTableCell(doc, blockFrom);
    if (!cell) {
      return { type: "skip", reason: "unsupportedBlock" };
    }
    const rectangle = {
      left: cell.leftColumnIndex,
      top: cell.topRowIndex,
      right: cell.rightColumnIndex,
      bottom: cell.bottomRowIndex,
    };
    if (rectangle.right - rectangle.left === 1 && rectangle.bottom - rectangle.top === 1) {
      return { type: "skip", reason: "noopOperation" };
    }
    return {
      type: "resolved",
      operation: {
        operation,
        from: cell.cellPosition,
        to: cell.cellEndPosition,
        blockFrom,
        blockTo,
        blockNode,
        tableCellSplit: {
          tablePosition: cell.tablePosition,
          rectangle,
        },
      },
    };
  }

  if (operation.type === "deleteBlock" || operation.type === "replaceBlock") {
    const range = getTextRangeFromCleanBlock(cleanBlock);
    if (!range) {
      const insertionPoint = cleanBlock.offsets.at(0);
      if (
        operation.type === "replaceBlock" &&
        currentText.length === 0 &&
        operation.text.length > 0 &&
        insertionPoint !== undefined
      ) {
        return {
          type: "resolved",
          operation: {
            operation,
            from: insertionPoint,
            to: insertionPoint,
            blockFrom,
            blockTo,
            blockNode,
          },
        };
      }
      return { type: "skip", reason: "unsupportedBlock" };
    }
    // The model occasionally emits replaceBlock with text identical
    // to the live block's clean text — usually as a side effect of
    // running through a "review every block" pass. Skip so the
    // panel doesn't show an empty redline (verified in dev-tools
    // trace where one such op surfaced as "Prodávající 3 →
    // Prodávající 3").
    if (operation.type === "replaceBlock" && operation.text === currentText) {
      return { type: "skip", reason: "noopOperation" };
    }
    return {
      type: "resolved",
      operation: {
        operation,
        from: range.from,
        to: range.to,
        blockFrom,
        blockTo,
        blockNode,
      },
    };
  }

  if (operation.type === "replaceInBlock" && operation.find === operation.replace) {
    return { type: "skip", reason: "noopOperation" };
  }

  const quote = getOperationQuote(operation);
  const range = resolveTextInCleanBlock(cleanBlock, quote || currentText);
  if (range.type !== "resolved") {
    return range;
  }

  return {
    type: "resolved",
    operation: {
      operation,
      from: range.from,
      to: range.to,
      blockFrom,
      blockTo,
      blockNode,
    },
  };
};

type OperationResolutionSkip = { type: "skip"; reason: FolioAIEditSkipReason };

const resolveTextInCleanBlock = (
  cleanBlock: { text: string; offsets: number[] },
  find: string,
):
  | { type: "resolved"; from: number; to: number }
  | { type: "skip"; reason: FolioAIEditSkipReason } => {
  if (find.length === 0) {
    return { type: "skip", reason: "emptyOperation" };
  }

  const { text, offsets } = cleanBlock;
  const firstIndex = text.indexOf(find);
  if (firstIndex === -1) {
    return { type: "skip", reason: "missingFind" };
  }
  if (text.includes(find, firstIndex + 1)) {
    return { type: "skip", reason: "ambiguousFind" };
  }

  const from = offsets[firstIndex];
  const to = offsets[firstIndex + find.length];
  if (from === undefined || to === undefined) {
    return { type: "skip", reason: "unsupportedBlock" };
  }

  return { type: "resolved", from, to };
};

const getTextRangeFromCleanBlock = (cleanBlock: {
  text: string;
  offsets: number[];
}): { from: number; to: number } | null => {
  if (cleanBlock.text.length === 0) {
    return null;
  }
  const from = cleanBlock.offsets[0];
  const to = cleanBlock.offsets[cleanBlock.text.length];
  if (from === undefined || to === undefined) {
    return null;
  }
  return { from, to };
};

const getOperationCommentText = (operation: FolioAIEditOperation): string | undefined =>
  "comment" in operation ? operation.comment?.text : undefined;

const getOperationQuote = (operation: FolioAIEditOperation): string | undefined => {
  if (operation.type === "replaceInBlock") {
    return operation.find;
  }
  if (operation.type === "commentOnBlock") {
    return operation.quote;
  }
  return undefined;
};
