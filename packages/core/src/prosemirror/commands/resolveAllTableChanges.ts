import { panic } from "better-result";
import type { Node as PMNode } from "prosemirror-model";
import { TableMap, type Rect } from "prosemirror-tables";
import { StepMap } from "prosemirror-transform";

import {
  restoreTableCellsWithParagraphPropertySources,
  transportTableCellsWithParagraphPropertySources,
} from "../../docx/paragraphPropertySource";
import type { TableCell } from "../../types/document";
import { expectTableAttrs, mergeTableAttrs } from "../attrs";
import { isTableCellRetainedInReviewView } from "../tableCellRevisionVisibility";
import { getTableCellMergeChange } from "../tableCellMergeRevision";
import { nodePropertyRevisionSites, propertyRevisionRecords } from "../revisionCarriers";
import {
  createRestoredTableCell,
  hasMatchingCollapsedTableCellMerge,
  tableCellContinuationCells,
  tableCellContinuationFromNode,
  tableCellContinuationPayload,
} from "./tableCellMergeResolution";
import { resolveAllNodePropertyChangeAttrs } from "./resolveNodePropertyChangeAttrs";

// TableMap.findCell scans the grid for each lookup. Index each rectangle once
// because whole-table resolution visits every cell.
const indexedTableCells = (map: TableMap) => {
  const rectangles = new Map<number, Rect>();
  for (let index = 0; index < map.map.length; index++) {
    const position = map.map[index];
    if (position === undefined || rectangles.has(position)) continue;
    const left = index % map.width;
    const top = Math.floor(index / map.width);
    let right = left + 1;
    let bottom = top + 1;
    while (right < map.width && map.map[index + right - left] === position) right++;
    while (bottom < map.height && map.map[index + (bottom - top) * map.width] === position)
      bottom++;
    rectangles.set(position, { left, top, right, bottom });
  }
  return (position: number): Rect =>
    rectangles.get(position) ?? panic("Missing table resolution rectangle", { position });
};

type ResolveMode = "accept" | "reject";

type ResolveAllTableChangesOptions = {
  /** The table before resolution; its topology determines row and cell positions. */
  table: PMNode;
  /** Descendants after the caller's inline, paragraph, and nested-table pass. */
  resolvedRows?: readonly PMNode[];
  mode: ResolveMode;
};

export type ResolvedTableChanges = {
  /** Null means every row was removed; the caller owns the enclosing schema fallback. */
  node: PMNode | null;
  changed: boolean;
  structural: boolean;
  failed: boolean;
  /** Final-table-relative ranges whose surviving paragraphs changed. */
  changedParagraphRanges: readonly { from: number; to: number }[];
  /** Source and final paragraph starts, each relative to its table node start. */
  paragraphOffsets: readonly { source: number; final: number }[];
  /** Table-node-relative source-to-final positions, preserving surviving cells. */
  positionMap: StepMap;
};

type CellPlacement = {
  node: PMNode;
  left: number;
  originRow: number;
  bottom: number;
  sourcePosition: number;
  mapSourcePosition: number;
  touched: boolean;
};

type CellMetadata = Pick<CellPlacement, "sourcePosition" | "mapSourcePosition" | "touched">;
type RowMetadata = { touched: boolean; cells: readonly CellMetadata[] };

const elementAt = <T>(items: readonly T[], index: number): T => {
  const item = items.at(index);
  return item === undefined ? panic("Missing table resolution element", { index }) : item;
};

type TableResolutionMetadataOptions = {
  sourceTable: PMNode;
  table: PMNode;
  rows: readonly RowMetadata[];
  tableTouched: boolean;
};

const tableResolutionMetadata = ({
  sourceTable,
  table,
  rows,
  tableTouched,
}: TableResolutionMetadataOptions) => {
  const changedParagraphRanges: { from: number; to: number }[] = [];
  const paragraphOffsets: { source: number; final: number }[] = [];
  const anchors: { source: number; final: number; size: number }[] = [];
  const finalCellStarts: number[] = [];
  if (tableTouched) {
    changedParagraphRanges.push({ from: 0, to: table.nodeSize });
  }
  let rowStart = 1;
  for (let rowIndex = 0; rowIndex < table.childCount; rowIndex++) {
    const row = table.child(rowIndex);
    const metadata = elementAt(rows, rowIndex);
    if (!tableTouched && metadata.touched) {
      changedParagraphRanges.push({ from: rowStart, to: rowStart + row.nodeSize });
    }
    let cellStart = rowStart + 1;
    for (let cellIndex = 0; cellIndex < row.childCount; cellIndex++) {
      const cell = row.child(cellIndex);
      const cellMetadata = elementAt(metadata.cells, cellIndex);
      finalCellStarts.push(cellStart);
      if (cellMetadata.mapSourcePosition >= 0) {
        anchors.push({
          source: cellMetadata.mapSourcePosition + 1,
          final: cellStart,
          size: cell.nodeSize,
        });
      }
      if (!tableTouched && !metadata.touched && cellMetadata.touched) {
        changedParagraphRanges.push({ from: cellStart, to: cellStart + cell.nodeSize });
      }
      if (cellMetadata.sourcePosition >= 0) {
        const sourceCellPosition = cellMetadata.sourcePosition;
        const finalCellPosition = cellStart;
        cell.descendants((node, offset) => {
          if (node.type.name === "table") return false;
          if (node.type.name !== "paragraph") return true;
          paragraphOffsets.push({
            source: sourceCellPosition + 2 + offset,
            final: finalCellPosition + 1 + offset,
          });
          return false;
        });
      }
      cellStart += cell.nodeSize;
    }
    rowStart += row.nodeSize;
  }
  const sourceCellStarts: number[] = [];
  let sourceRowStart = 1;
  for (let rowIndex = 0; rowIndex < sourceTable.childCount; rowIndex++) {
    const row = sourceTable.child(rowIndex);
    let cellStart = sourceRowStart + 1;
    for (let cellIndex = 0; cellIndex < row.childCount; cellIndex++) {
      sourceCellStarts.push(cellStart);
      cellStart += row.child(cellIndex).nodeSize;
    }
    sourceRowStart += row.nodeSize;
  }
  const ranges: number[] = [];
  let sourceCursor = 0;
  let finalCursor = 0;
  let sourceCellIndex = 0;
  let finalCellIndex = 0;
  for (const anchor of [
    ...anchors,
    { source: sourceTable.nodeSize, final: table.nodeSize, size: 0 },
  ]) {
    if (anchor.source < sourceCursor || anchor.final < finalCursor) {
      continue;
    }
    while (
      sourceCellStarts.at(sourceCellIndex) !== undefined &&
      elementAt(sourceCellStarts, sourceCellIndex) < sourceCursor
    ) {
      sourceCellIndex++;
    }
    while (
      finalCellStarts.at(finalCellIndex) !== undefined &&
      elementAt(finalCellStarts, finalCellIndex) < finalCursor
    ) {
      finalCellIndex++;
    }
    const sourceGap = anchor.source - sourceCursor;
    const finalGap = anchor.final - finalCursor;
    const hasUnmatchedCell =
      (sourceCellStarts.at(sourceCellIndex) ?? Infinity) < anchor.source ||
      (finalCellStarts.at(finalCellIndex) ?? Infinity) < anchor.final;
    if (sourceGap !== finalGap || hasUnmatchedCell) {
      ranges.push(sourceCursor, sourceGap, finalGap);
    }
    sourceCursor = anchor.source + anchor.size;
    finalCursor = anchor.final + anchor.size;
  }
  return { changedParagraphRanges, paragraphOffsets, positionMap: new StepMap(ranges) };
};

const hasNodePropertyRevision = (node: PMNode): boolean =>
  nodePropertyRevisionSites(node.type.name).some(
    (site) => site.resolution === "node-attrs" && propertyRevisionRecords(node, site).length > 0,
  );

const rowRevisionKind = (row: PMNode): "trIns" | "trDel" | null => {
  for (const kind of ["trIns", "trDel"] as const) {
    const marker = row.attrs[kind];
    if (
      typeof marker === "object" &&
      marker !== null &&
      "revisionId" in marker &&
      typeof marker.revisionId === "number"
    ) {
      return kind;
    }
  }
  return null;
};

type CellMarker = {
  kind: "ins" | "del" | "merge";
  info: { revisionId: number };
  verticalMergeOriginal?: "continue" | "rest";
};

const cellRevisionMarker = (cell: PMNode): CellMarker | null => {
  const marker = cell.attrs["cellMarker"];
  if (
    typeof marker !== "object" ||
    marker === null ||
    !("kind" in marker) ||
    (marker.kind !== "ins" && marker.kind !== "del" && marker.kind !== "merge") ||
    !("info" in marker) ||
    typeof marker.info !== "object" ||
    marker.info === null ||
    !("revisionId" in marker.info) ||
    typeof marker.info.revisionId !== "number"
  ) {
    return null;
  }
  const verticalMergeOriginal =
    "verticalMergeOriginal" in marker &&
    (marker.verticalMergeOriginal === "continue" || marker.verticalMergeOriginal === "rest")
      ? marker.verticalMergeOriginal
      : undefined;
  return {
    kind: marker.kind,
    info: { revisionId: marker.info.revisionId },
    ...(verticalMergeOriginal ? { verticalMergeOriginal } : {}),
  };
};

/**
 * Resolve a table in one grid pass for row revisions and ordinary cell
 * membership revisions. The grid records each original cell rectangle once;
 * retaining rows reduces its rowspan, while a cell whose starting row is
 * removed moves to the next retained row inside that rectangle. This is the
 * same topology operation as prosemirror-tables' removeRow.
 */
export const resolveAllTableChanges = ({
  table,
  resolvedRows,
  mode,
}: ResolveAllTableChangesOptions): ResolvedTableChanges => {
  const sourceRows =
    resolvedRows ?? Array.from({ length: table.childCount }, (_, i) => table.child(i));
  const tableWithResolvedChildren = table.type.create(table.attrs, sourceRows, table.marks);
  const tablePropertyTouched = hasNodePropertyRevision(tableWithResolvedChildren);
  const rowPropertyTouched = sourceRows.map(hasNodePropertyRevision);
  const propertyResolvedTable = resolveTablePropertyAttrs(tableWithResolvedChildren, mode);
  const rows = Array.from({ length: propertyResolvedTable.childCount }, (_, i) =>
    propertyResolvedTable.child(i),
  );
  const rowKinds = rows.map(rowRevisionKind);
  const keepRows = rowKinds.map(
    (kind) => kind === null || (kind === "trIns") === (mode === "accept"),
  );
  const hasRowRevision = rowKinds.some((kind) => kind !== null);
  const hasMerge = rows.some((row) => {
    for (let i = 0; i < row.childCount; i++) {
      const cell = row.child(i);
      if (
        cellRevisionMarker(cell)?.kind === "merge" ||
        hasMatchingCollapsedTableCellMerge(cell, null)
      ) {
        return true;
      }
    }
    return false;
  });
  const map = TableMap.get(propertyResolvedTable);
  const cellRectangle = indexedTableCells(map);
  const cells: CellPlacement[] = [];
  let rowOffset = 0;
  let hasCellRevision = false;
  for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
    const row = elementAt(rows, rowIndex);
    let cellOffset = rowOffset + 1;
    for (let cellIndex = 0; cellIndex < row.childCount; cellIndex++) {
      const cell = row.child(cellIndex);
      const rectangle = cellRectangle(cellOffset);
      const marker = cellRevisionMarker(cell);
      const membership = marker?.kind === "ins" || marker?.kind === "del" ? marker : null;
      if (membership !== null) {
        hasCellRevision = true;
      }
      cells.push({
        node: cell,
        left: rectangle.left,
        originRow: rowIndex,
        bottom: rectangle.bottom,
        sourcePosition: cellOffset,
        mapSourcePosition: cellOffset,
        touched: hasNodePropertyRevision(elementAt(sourceRows, rowIndex).child(cellIndex)),
      });
      cellOffset += cell.nodeSize;
    }
    rowOffset += row.nodeSize;
  }

  if (!keepRows.some(Boolean)) {
    return {
      node: null,
      changed: true,
      structural: hasRowRevision || hasCellRevision,
      failed: false,
      changedParagraphRanges: [],
      paragraphOffsets: [],
      positionMap: new StepMap([0, tableWithResolvedChildren.nodeSize, 0]),
    };
  }

  const keptBefore = [0];
  for (const keep of keepRows) {
    keptBefore.push((keptBefore.at(-1) ?? 0) + Number(keep));
  }
  const nextKeptRow = rows.map(() => rows.length);
  let next = rows.length;
  for (let index = rows.length - 1; index >= 0; index--) {
    if (elementAt(keepRows, index)) next = index;
    nextKeptRow[index] = next;
  }
  const placements = rows.map((): (CellPlacement | undefined)[] => []);
  for (const cell of cells) {
    const destination = elementAt(nextKeptRow, cell.originRow);
    if (destination >= cell.bottom) {
      continue;
    }
    const rowspan = elementAt(keptBefore, cell.bottom) - elementAt(keptBefore, destination);
    elementAt(placements, destination)[cell.left] = {
      ...cell,
      mapSourcePosition: destination === cell.originRow ? cell.mapSourcePosition : -1,
      touched: cell.touched || rowspan !== cell.node.attrs["rowspan"],
      node:
        rowspan === cell.node.attrs["rowspan"]
          ? cell.node
          : cell.node.type.create(
              { ...cell.node.attrs, rowspan },
              cell.node.content,
              cell.node.marks,
            ),
    };
  }

  const nextRows: PMNode[] = [];
  const rowMetadata: RowMetadata[] = [];
  for (let index = 0; index < rows.length; index++) {
    if (!elementAt(keepRows, index)) {
      continue;
    }
    const row = elementAt(rows, index);
    const rowCells = elementAt(placements, index).flatMap((placement) =>
      placement === undefined ? [] : [placement.node],
    );
    const kind = elementAt(rowKinds, index);
    const attrs = kind ? { ...row.attrs, [kind]: null } : row.attrs;
    nextRows.push(row.type.create(attrs, rowCells, row.marks));
    rowMetadata.push({
      touched: elementAt(rowPropertyTouched, index) || kind !== null,
      cells: elementAt(placements, index).flatMap((placement) =>
        placement === undefined ? [] : [placement],
      ),
    });
  }
  const nextTable = propertyResolvedTable.type.create(
    propertyResolvedTable.attrs,
    nextRows,
    propertyResolvedTable.marks,
  );
  if (hasMerge || hasCellRevision) {
    return resolvePureTableMerges({
      table: nextTable,
      mode,
      rowMetadata,
      tablePropertyTouched,
      sourceTable: tableWithResolvedChildren,
    });
  }
  const structural =
    hasRowRevision || hasCellRevision || hasTablePropertyRevision(tableWithResolvedChildren);
  return {
    node: nextTable,
    changed: structural || !nextTable.eq(table),
    structural,
    failed: false,
    ...tableResolutionMetadata({
      table: nextTable,
      rows: rowMetadata,
      tableTouched: tablePropertyTouched,
      sourceTable: tableWithResolvedChildren,
    }),
  };
};

const resolveTablePropertyAttrs = (table: PMNode, mode: ResolveMode): PMNode => {
  const rows: PMNode[] = [];
  for (let rowIndex = 0; rowIndex < table.childCount; rowIndex++) {
    const row = table.child(rowIndex);
    const cells: PMNode[] = [];
    for (let cellIndex = 0; cellIndex < row.childCount; cellIndex++) {
      cells.push(resolveAllNodePropertyChangeAttrs(row.child(cellIndex), mode));
    }
    rows.push(
      resolveAllNodePropertyChangeAttrs(row.type.create(row.attrs, cells, row.marks), mode),
    );
  }
  return resolveAllNodePropertyChangeAttrs(table.type.create(table.attrs, rows, table.marks), mode);
};

const hasTablePropertyRevision = (table: PMNode): boolean => {
  if (table.attrs["tblPrChange"] != null) {
    return true;
  }
  for (let rowIndex = 0; rowIndex < table.childCount; rowIndex++) {
    const row = table.child(rowIndex);
    if (row.attrs["trPrChange"] != null || row.attrs["tblPrExChange"] != null) {
      return true;
    }
    for (let cellIndex = 0; cellIndex < row.childCount; cellIndex++) {
      if (row.child(cellIndex).attrs["tcPrChange"] != null) {
        return true;
      }
    }
  }
  return false;
};

type MutableCell = {
  node: PMNode;
  position: number;
  left: number;
  row: number;
  sourcePosition: number;
  mapSourcePosition: number;
  touched: boolean;
  continuation: ContinuationSequence | null;
};

type ContinuationLink = { cell: TableCell; next: ContinuationLink | null };
type ContinuationSequence = { head: ContinuationLink | null; tail: ContinuationLink | null };

const appendContinuation = (sequence: ContinuationSequence, cell: TableCell): void => {
  const link = { cell, next: null };
  if (sequence.tail) sequence.tail.next = link;
  else sequence.head = link;
  sequence.tail = link;
};

const continuationSequence = (cells: readonly TableCell[]): ContinuationSequence => {
  const sequence: ContinuationSequence = { head: null, tail: null };
  for (const cell of cells) appendContinuation(sequence, cell);
  return sequence;
};

const appendContinuationSequence = (
  destination: ContinuationSequence,
  source: ContinuationSequence,
): void => {
  if (!source.head) return;
  if (destination.tail) destination.tail.next = source.head;
  else destination.head = source.head;
  destination.tail = source.tail;
};

const materializeContinuation = (entry: MutableCell): void => {
  if (!entry.continuation) return;
  const cells: TableCell[] = [];
  for (let link = entry.continuation.head; link !== null; link = link.next) {
    cells.push(link.cell);
  }
  entry.node = entry.node.type.create(
    {
      ...entry.node.attrs,
      _docxVMergeContinuationCells: transportTableCellsWithParagraphPropertySources(cells),
    },
    entry.node.content,
    entry.node.marks,
  );
  entry.continuation = null;
};

type ResolvePureTableMergesOptions = {
  sourceTable: PMNode;
  table: PMNode;
  mode: ResolveMode;
  rowMetadata: readonly RowMetadata[];
  tablePropertyTouched: boolean;
};

const failedTableResolution = (table: PMNode): ResolvedTableChanges => ({
  node: table,
  changed: false,
  structural: false,
  failed: true,
  changedParagraphRanges: [],
  paragraphOffsets: [],
  positionMap: StepMap.empty,
});

/** Resolve all cell edits in reverse order against one grid and mutable slots. */
const resolvePureTableMerges = ({
  table,
  mode,
  rowMetadata,
  tablePropertyTouched,
  sourceTable,
}: ResolvePureTableMergesOptions): ResolvedTableChanges => {
  const map = TableMap.get(table);
  const cellRectangle = indexedTableCells(map);
  const rows = Array.from({ length: table.childCount }, (_, index) => table.child(index));
  const rowCells = rows.map(() => new Map<number, MutableCell>());
  const rowOccupancy = rows.map(() => 0);
  const activeRows = rows.map(() => true);
  const widthCounts = new Map<number, number>();
  const authoredWidths = expectTableAttrs(table).columnWidths;
  const columnWidths = authoredWidths?.length === map.width ? [...authoredWidths] : undefined;
  let currentWidth = map.width;
  let removedCell = false;
  const deletedRows = new Set<number>();
  const byPosition = new Map<number, MutableCell>();
  const ordered: MutableCell[] = [];
  let rowOffset = 0;
  for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
    const row = elementAt(rows, rowIndex);
    let cellOffset = rowOffset + 1;
    for (let cellIndex = 0; cellIndex < row.childCount; cellIndex++) {
      const node = row.child(cellIndex);
      const rectangle = cellRectangle(cellOffset);
      const left = rectangle.left;
      for (let coveredRow = rectangle.top; coveredRow < rectangle.bottom; coveredRow++) {
        rowOccupancy[coveredRow] =
          elementAt(rowOccupancy, coveredRow) + (rectangle.right - rectangle.left);
      }
      const initial = elementAt(elementAt(rowMetadata, rowIndex).cells, cellIndex);
      const entry: MutableCell = {
        node,
        position: cellOffset,
        left,
        row: rowIndex,
        sourcePosition: initial.sourcePosition,
        mapSourcePosition: initial.mapSourcePosition,
        touched: initial.touched,
        continuation: null,
      };
      elementAt(rowCells, rowIndex).set(left, entry);
      byPosition.set(cellOffset, entry);
      ordered.push(entry);
      cellOffset += node.nodeSize;
    }
    rowOffset += row.nodeSize;
  }
  for (const width of rowOccupancy) {
    widthCounts.set(width, (widthCounts.get(width) ?? 0) + 1);
  }

  const changeRowWidth = (rowIndex: number, nextWidth: number): void => {
    if (!elementAt(activeRows, rowIndex)) return;
    const previous = elementAt(rowOccupancy, rowIndex);
    widthCounts.set(previous, (widthCounts.get(previous) ?? 0) - 1);
    rowOccupancy[rowIndex] = nextWidth;
    widthCounts.set(nextWidth, (widthCounts.get(nextWidth) ?? 0) + 1);
  };

  for (let index = ordered.length - 1; index >= 0; index--) {
    const entry = elementAt(ordered, index);
    const marker = cellRevisionMarker(entry.node);
    if (marker && marker.kind !== "merge") {
      const view = mode === "accept" ? "final" : "original";
      if (isTableCellRetainedInReviewView(marker.kind, view)) {
        entry.touched = true;
        entry.node = entry.node.type.create(
          { ...entry.node.attrs, cellMarker: null },
          entry.node.content,
          entry.node.marks,
        );
      } else {
        removedCell = true;
        const previousWidth = currentWidth;
        const rectangle = cellRectangle(entry.position);
        const colspan = rectangle.right - rectangle.left;
        for (let coveredRow = rectangle.top; coveredRow < rectangle.bottom; coveredRow++) {
          changeRowWidth(coveredRow, elementAt(rowOccupancy, coveredRow) - colspan);
        }
        elementAt(rowCells, entry.row).delete(entry.left);
        if (elementAt(rowCells, entry.row).size === 0) {
          deletedRows.add(entry.row);
          activeRows[entry.row] = false;
          const rowWidth = elementAt(rowOccupancy, entry.row);
          widthCounts.set(rowWidth, (widthCounts.get(rowWidth) ?? 0) - 1);
        }
        while (currentWidth > 0 && (widthCounts.get(currentWidth) ?? 0) === 0) {
          currentWidth--;
        }
        if (columnWidths && currentWidth < previousWidth) {
          columnWidths.splice(entry.left, previousWidth - currentWidth);
        }
      }
    } else if (marker?.kind === "merge") {
      if (mode === "accept") {
        entry.touched = true;
        entry.node = entry.node.type.create(
          { ...entry.node.attrs, cellMarker: null },
          entry.node.content,
          entry.node.marks,
        );
      } else if ((marker.verticalMergeOriginal ?? "rest") === "rest") {
        entry.touched = true;
        const original = entry.node.attrs["_originalFormatting"];
        let formatting: unknown = original;
        if (typeof original === "object" && original !== null) {
          const copy: Record<string, unknown> = { ...original };
          delete copy["vMerge"];
          formatting = copy;
        }
        entry.node = entry.node.type.create(
          { ...entry.node.attrs, cellMarker: null, _originalFormatting: formatting },
          entry.node.content,
          entry.node.marks,
        );
      } else {
        const rectangle = cellRectangle(entry.position);
        if (rectangle.top === 0) {
          return failedTableResolution(table);
        }
        const abovePosition = map.map[(rectangle.top - 1) * map.width + rectangle.left];
        if (abovePosition === undefined) {
          return failedTableResolution(table);
        }
        const aboveRectangle = cellRectangle(abovePosition);
        const above = byPosition.get(abovePosition);
        if (
          !above ||
          !elementAt(rowCells, above.row).has(above.left) ||
          aboveRectangle.left !== rectangle.left ||
          aboveRectangle.right !== rectangle.right ||
          aboveRectangle.bottom !== rectangle.top
        ) {
          return failedTableResolution(table);
        }
        const aboveRowspan = above.node.attrs["rowspan"];
        const cellRowspan = entry.node.attrs["rowspan"];
        if (
          typeof aboveRowspan !== "number" ||
          typeof cellRowspan !== "number" ||
          aboveRowspan < 1 ||
          cellRowspan < 1
        ) {
          return failedTableResolution(table);
        }
        above.continuation ??= continuationSequence(
          tableCellContinuationCells(above.node, aboveRowspan),
        );
        // Legacy's topology replacement does not mark the surviving anchor
        // paragraph in the change tracker.
        appendContinuation(above.continuation, tableCellContinuationFromNode(entry.node));
        const nested =
          entry.continuation ??
          continuationSequence(tableCellContinuationPayload(entry.node)?.cells ?? []);
        appendContinuationSequence(above.continuation, nested);
        above.node = above.node.type.create(
          {
            ...above.node.attrs,
            rowspan: aboveRowspan + cellRowspan,
          },
          above.node.content,
          above.node.marks,
        );
        elementAt(rowCells, entry.row).delete(entry.left);
      }
    }

    if (!elementAt(rowCells, entry.row).has(entry.left)) {
      continue;
    }
    const originalPayload = tableCellContinuationPayload(entry.node);
    if (!originalPayload?.cells.some((cell) => getTableCellMergeChange(cell) !== null)) continue;
    materializeContinuation(entry);
    const payload = tableCellContinuationPayload(entry.node);
    if (!payload) return failedTableResolution(table);
    const stored = payload.cells;
    const splitIndices: number[] = [];
    let matched = false;
    const nextCells = stored.map((source, sourceIndex) => {
      const change = getTableCellMergeChange(source);
      if (!change) return source;
      matched = true;
      if (
        mode === "reject" &&
        change.verticalMerge === "continue" &&
        (change.verticalMergeOriginal ?? "rest") === "rest"
      ) {
        splitIndices.push(sourceIndex);
      }
      const cleaned = { ...source };
      delete cleaned.structuralChange;
      return cleaned;
    });
    if (!matched) continue;
    if (mode === "accept" || splitIndices.length === 0) {
      entry.touched = true;
      entry.node = entry.node.type.create(
        { ...entry.node.attrs, _docxVMergeContinuationCells: nextCells },
        entry.node.content,
        entry.node.marks,
      );
      continue;
    }
    const firstSplitIndex = splitIndices.at(0);
    if (
      firstSplitIndex === undefined ||
      splitIndices.length !== stored.length - firstSplitIndex ||
      splitIndices.some((splitIndex, offset) => splitIndex !== firstSplitIndex + offset)
    ) {
      return failedTableResolution(table);
    }
    const rectangle = cellRectangle(entry.position);
    const rowspan = entry.node.attrs["rowspan"];
    if (
      typeof rowspan !== "number" ||
      rowspan !== stored.length + 1 ||
      rectangle.bottom - rectangle.top !== rowspan
    ) {
      return failedTableResolution(table);
    }
    const restored = restoreTableCellsWithParagraphPropertySources(payload);
    for (const splitIndex of splitIndices) {
      const source = restored[splitIndex];
      const restoredCell = source && createRestoredTableCell(entry.node, source);
      const rowIndex = rectangle.top + splitIndex + 1;
      if (!restoredCell || elementAt(rowCells, rowIndex).has(rectangle.left)) {
        return failedTableResolution(table);
      }
      elementAt(rowCells, rowIndex).set(rectangle.left, {
        node: restoredCell,
        position: -1,
        left: rectangle.left,
        row: rowIndex,
        sourcePosition: -1,
        mapSourcePosition: -1,
        touched: true,
        continuation: null,
      });
    }
    const remaining = nextCells.slice(0, firstSplitIndex);
    entry.node = entry.node.type.create(
      {
        ...entry.node.attrs,
        rowspan: firstSplitIndex + 1,
        _docxVMergeContinuationCells: remaining.length > 0 ? remaining : null,
      },
      entry.node.content,
      entry.node.marks,
    );
  }

  if (deletedRows.size === rows.length) {
    return {
      node: null,
      changed: true,
      structural: true,
      failed: false,
      changedParagraphRanges: [],
      paragraphOffsets: [],
      positionMap: new StepMap([0, sourceTable.nodeSize, 0]),
    };
  }
  const keptBefore = [0];
  for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
    keptBefore.push((keptBefore.at(-1) ?? 0) + Number(!deletedRows.has(rowIndex)));
  }
  const compacted = rows.map(() => new Map<number, MutableCell>());
  for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
    for (const entry of elementAt(rowCells, rowIndex).values()) {
      materializeContinuation(entry);
      const originalRowspan = entry.node.attrs["rowspan"];
      if (typeof originalRowspan !== "number") continue;
      const bottom = Math.min(rowIndex + originalRowspan, rows.length);
      let destination = rowIndex;
      while (destination < bottom && deletedRows.has(destination)) destination++;
      if (destination === bottom) continue;
      const rowspan = elementAt(keptBefore, bottom) - elementAt(keptBefore, destination);
      const node =
        rowspan === originalRowspan
          ? entry.node
          : entry.node.type.create(
              { ...entry.node.attrs, rowspan },
              entry.node.content,
              entry.node.marks,
            );
      elementAt(compacted, destination).set(entry.left, {
        ...entry,
        node,
        row: destination,
        touched: entry.touched || rowspan !== originalRowspan,
      });
    }
  }
  const nextRowMetadata: RowMetadata[] = [];
  const nextRows = rows.flatMap((row, rowIndex) => {
    if (deletedRows.has(rowIndex)) return [];
    const cells: PMNode[] = [];
    const cellMetadata: CellMetadata[] = [];
    for (let column = 0; column < map.width; column++) {
      const entry = elementAt(compacted, rowIndex).get(column);
      if (entry) {
        cells.push(entry.node);
        cellMetadata.push(entry);
      }
    }
    nextRowMetadata.push({
      touched: elementAt(rowMetadata, rowIndex).touched,
      cells: cellMetadata,
    });
    return [row.type.create(row.attrs, cells, row.marks)];
  });
  let nextTable = table.type.create(table.attrs, nextRows, table.marks);
  const nextWidth = TableMap.get(nextTable).width;
  const gridChanged = removedCell && nextWidth < map.width;
  if (gridChanged) {
    const formatting = expectTableAttrs(nextTable)._originalFormatting;
    const originalFormatting = formatting ? { ...formatting } : undefined;
    if (originalFormatting) delete originalFormatting.gridSourceXml;
    nextTable = nextTable.type.create(
      mergeTableAttrs(nextTable, {
        columnWidths: columnWidths?.length === nextWidth ? columnWidths : undefined,
        _originalFormatting: originalFormatting,
      }),
      nextTable.content,
      nextTable.marks,
    );
  }
  return {
    node: nextTable,
    changed: true,
    structural: true,
    failed: false,
    ...tableResolutionMetadata({
      table: nextTable,
      rows: nextRowMetadata,
      tableTouched: tablePropertyTouched || gridChanged,
      sourceTable,
    }),
  };
};
