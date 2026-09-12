import { panic } from "better-result";
import type { Node as PMNode } from "prosemirror-model";
import type { Transaction } from "prosemirror-state";
import { TableMap } from "prosemirror-tables";

import {
  markTableRowContent,
  type MarkTableRowContentOptions,
} from "../../ai-edits/table-row-column-mutations";
import {
  tableFromTemplate,
  tableTemplateCanCrossPackageLosslessly,
} from "../../ai-edits/table-template";
import { storyTablesOf } from "../../ai-edits/snapshot";
import { expectTableAttrs, mergeTableAttrs } from "../../prosemirror/attrs";
import { markStructuralChange } from "../../prosemirror/extensions/features/ParagraphChangeTrackerExtension";
import { canonicalJson } from "../../utils/canonicalJson";
import {
  resolvedDocxStoryComparisonPayload,
  resolvedDocxTableComponentOperands,
  resolvedDocxTableFormatOperandPayload,
  resolvedDocxTableGeometryPairings,
  resolvedDocxTableStructureOperandPayload,
  type ResolvedDocxStoryComparison,
  type ResolvedDocxTableComponent,
  type ResolvedDocxTableFormatOperand,
  type ResolvedDocxTableStructureOperand,
  type ResolvedDocxTableStructureOperandPayload,
} from "./resolved-docx-story-comparison";
import {
  resolvedDocxOperationSnapshot,
  findResolvedDocxTableNode,
  resolvedDocxSourceOperandBlock,
  type ResolvedDocxSourceOperand,
  type ResolvedDocxStorySnapshot,
} from "./resolved-docx-story-snapshot";
import { FolioStableBlockResolver } from "./stable-block-resolution";
import {
  executeTableGeometryProgram,
  tableGeometryProgramTableGridTransitions,
  type TableGeometryExecutionIssue,
  type TableGeometryExecutionReceipt,
  type TableGeometryPairing,
  type TableGeometryProgram,
  type TableGeometryUnsupportedIssue,
} from "./table-geometry-program";

type ResolvedDocxTableOperand = ResolvedDocxTableStructureOperand | ResolvedDocxTableFormatOperand;

type ResolvedDocxTableStructureOperation = {
  [Type in ResolvedDocxTableStructureOperand["type"]]: {
    readonly type: Type;
    readonly operand: Extract<ResolvedDocxTableStructureOperand, { readonly type: Type }>;
    readonly payload: Extract<ResolvedDocxTableStructureOperandPayload, { readonly type: Type }>;
  };
}[ResolvedDocxTableStructureOperand["type"]];

const resolveTableStructureOperation = (
  operand: ResolvedDocxTableStructureOperand,
  comparison: ResolvedDocxStoryComparison,
): ResolvedDocxTableStructureOperation => {
  const payload = resolvedDocxTableStructureOperandPayload(operand, comparison);
  switch (operand.type) {
    case "insertTable":
      if (payload.type !== "insertTable") {
        return panic("A table insertion lost its canonical payload");
      }
      return Object.freeze({ type: "insertTable", operand, payload });
    case "deleteTable":
      if (payload.type !== "deleteTable") {
        return panic("A table deletion lost its canonical payload");
      }
      return Object.freeze({ type: "deleteTable", operand, payload });
    case "replaceTable":
      if (payload.type !== "replaceTable") {
        return panic("A table replacement lost its canonical payload");
      }
      return Object.freeze({ type: "replaceTable", operand, payload });
    case "insertTableRow":
      if (payload.type !== "insertTableRow") {
        return panic("A table-row insertion lost its canonical payload");
      }
      return Object.freeze({ type: "insertTableRow", operand, payload });
    case "deleteTableRow":
      if (payload.type !== "deleteTableRow") {
        return panic("A table-row deletion lost its canonical payload");
      }
      return Object.freeze({ type: "deleteTableRow", operand, payload });
    case "insertTableColumn":
      if (payload.type !== "insertTableColumn") {
        return panic("A table-column insertion lost its canonical payload");
      }
      return Object.freeze({ type: "insertTableColumn", operand, payload });
    case "deleteTableColumn":
      if (payload.type !== "deleteTableColumn") {
        return panic("A table-column deletion lost its canonical payload");
      }
      return Object.freeze({ type: "deleteTableColumn", operand, payload });
    default: {
      const unreachable: never = operand;
      return panic("Unhandled table-structure operand", { operand: unreachable });
    }
  }
};

const TABLE_ROLE = "table";
const ROW_ROLE = "row";
const CELL_ROLES = new Set(["cell", "header_cell"]);

export type TableStructurePreflightLimits = {
  readonly maxOperands: number;
  readonly maxStructuralEdits: number;
  readonly maxDocumentUnits: number;
  readonly maxTargetTemplateUnits: number;
  readonly maxTables: number;
};

export const DEFAULT_TABLE_STRUCTURE_PREFLIGHT_LIMITS = Object.freeze({
  maxOperands: 10_000,
  maxStructuralEdits: 10_000,
  maxDocumentUnits: 4_194_304,
  maxTargetTemplateUnits: 4_194_304,
  maxTables: 10_000,
} as const satisfies TableStructurePreflightLimits);

type TableStructureLimit = keyof TableStructurePreflightLimits;

export type TableStructureUnsupportedIssue =
  | {
      readonly reason: "invalid-limit";
      readonly limit: TableStructureLimit;
      readonly actual: number;
    }
  | {
      readonly reason: "limit-exceeded";
      readonly limit: TableStructureLimit;
      readonly maximum: number;
      readonly actual: number;
    }
  | { readonly reason: "duplicate-operand" }
  | { readonly reason: "duplicate-geometry-operand" }
  | { readonly reason: "missing-source-table" }
  | { readonly reason: "source-table-mismatch" }
  | { readonly reason: "missing-target-table" }
  | { readonly reason: "unexpected-node-role"; readonly side: "source" | "target" }
  | { readonly reason: "pending-structural-revision"; readonly side: "source" | "target" }
  | { readonly reason: "unprojected-table-structure"; readonly side: "source" | "target" }
  | { readonly reason: "package-bound-template"; readonly side: "target" }
  | { readonly reason: "invalid-insertion-boundary" }
  | { readonly reason: "invalid-structural-coordinate" }
  | { readonly reason: "duplicate-structural-coordinate" }
  | { readonly reason: "unrepresentable-span"; readonly side: "source" | "target" }
  | { readonly reason: "non-reconstructable-structure" }
  | { readonly reason: "mixed-table-structure-axes" }
  | { readonly reason: "overlapping-table-obligations" }
  | { readonly reason: "duplicate-source-table" }
  | {
      readonly reason:
        | "terminal-carrier-missing"
        | "terminal-carrier-not-empty"
        | "terminal-carrier-not-final"
        | "terminal-source-not-body-peer";
    }
  | {
      readonly reason: "unrepresentable-table-geometry";
      readonly issue: TableGeometryUnsupportedIssue;
    };

export type TableStructurePreflightResult =
  | { readonly status: "ready"; readonly program: PreparedTableStructureProgram }
  | { readonly status: "unsupported"; readonly issue: TableStructureUnsupportedIssue };

export type TableStructureComponentPreflightResult = {
  readonly component: ResolvedDocxTableComponent;
  readonly result: TableStructurePreflightResult;
};

export type TableStructureComponentsPreflightResult =
  | {
      readonly status: "ready";
      readonly components: readonly TableStructureComponentPreflightResult[];
    }
  | { readonly status: "unsupported"; readonly issue: TableStructureUnsupportedIssue };

type SourceTable = {
  readonly index: number;
  readonly position: number;
  readonly expected: PMNode;
  readonly structuralState: string;
};

type InsertionBoundary = {
  readonly position: number;
  readonly depth: number;
  readonly parentTypeName: string;
};

type TerminalCarrier = {
  readonly position: number;
  readonly expected: PMNode;
};

type RowEdit =
  | {
      readonly type: "insert";
      readonly operand: Extract<
        ResolvedDocxTableStructureOperand,
        { readonly type: "insertTableRow" }
      >;
      readonly baseBoundaryIndex: number;
      readonly targetRowIndex: number;
      readonly targetRow: PMNode;
    }
  | {
      readonly type: "delete";
      readonly operand: Extract<
        ResolvedDocxTableStructureOperand,
        { readonly type: "deleteTableRow" }
      >;
      readonly baseRowIndex: number;
    };

type ColumnEdit =
  | {
      readonly type: "insert";
      readonly operand: Extract<
        ResolvedDocxTableStructureOperand,
        { readonly type: "insertTableColumn" }
      >;
      readonly baseBoundaryIndex: number;
      readonly targetColumnIndex: number;
      readonly targetCells: readonly PMNode[];
    }
  | {
      readonly type: "delete";
      readonly operand: Extract<
        ResolvedDocxTableStructureOperand,
        { readonly type: "deleteTableColumn" }
      >;
      readonly baseColumnIndex: number;
    };

type OwnedInstruction =
  | {
      readonly type: "insertTable";
      readonly operand: Extract<
        ResolvedDocxTableStructureOperand,
        { readonly type: "insertTable" }
      >;
      readonly boundary: InsertionBoundary;
      readonly target: PMNode;
      readonly terminalCarrier?: TerminalCarrier;
    }
  | {
      readonly type: "deleteTable";
      readonly operand: Extract<
        ResolvedDocxTableStructureOperand,
        { readonly type: "deleteTable" }
      >;
      readonly source: SourceTable;
    }
  | {
      readonly type: "replaceTable";
      readonly operand: Extract<
        ResolvedDocxTableStructureOperand,
        { readonly type: "replaceTable" }
      >;
      readonly source: SourceTable;
      readonly target: PMNode;
      readonly terminalCarrier?: TerminalCarrier;
    }
  | {
      readonly type: "editTableRows";
      readonly source: SourceTable;
      readonly target: PMNode;
      readonly edits: readonly RowEdit[];
    }
  | {
      readonly type: "editTableColumns";
      readonly source: SourceTable;
      readonly target: PMNode;
      readonly edits: readonly ColumnEdit[];
      readonly trackedColumnWidths?: readonly number[];
    };

type OwnedTableInsertion = Extract<OwnedInstruction, { readonly type: "insertTable" }>;
type OwnedWholeTableInstruction = Extract<
  OwnedInstruction,
  { readonly type: "deleteTable" | "replaceTable" }
>;

const PREPARED_TABLE_STRUCTURE_PROGRAM: unique symbol = Symbol("prepared-table-structure-program");
const PREPARED_TABLE_STRUCTURE_TASK: unique symbol = Symbol("prepared-table-structure-task");

export type PreparedTableStructureTask = {
  readonly [PREPARED_TABLE_STRUCTURE_TASK]: true;
  readonly schedule:
    | {
        readonly phase: "source";
        readonly position: number;
        readonly from: number;
        readonly to: number;
      }
    | { readonly phase: "insertion"; readonly position: number };
  readonly operands: readonly ResolvedDocxTableStructureOperand[];
};

export type PreparedTableStructureProgram = {
  readonly [PREPARED_TABLE_STRUCTURE_PROGRAM]: true;
  readonly tasks: readonly PreparedTableStructureTask[];
  readonly operands: readonly ResolvedDocxTableOperand[];
  readonly geometryOperand?: ResolvedDocxTableFormatOperand;
};

type OwnedProgram = {
  readonly sourceDocument: PMNode;
  readonly instructionsByTask: ReadonlyMap<PreparedTableStructureTask, OwnedInstruction>;
  readonly geometry: {
    readonly operand: ResolvedDocxTableFormatOperand;
    readonly program: TableGeometryProgram;
  } | null;
  readonly executedTasks: Set<PreparedTableStructureTask>;
  geometryExecuted: boolean;
};

const ownedPrograms = new WeakMap<PreparedTableStructureProgram, OwnedProgram>();

const unsupported = (
  issue: TableStructureUnsupportedIssue,
): Extract<TableStructurePreflightResult, { readonly status: "unsupported" }> =>
  Object.freeze({ status: "unsupported", issue: Object.freeze(issue) });

const isTable = (node: PMNode): boolean => node.type.spec["tableRole"] === TABLE_ROLE;
const isRow = (node: PMNode): boolean => node.type.spec["tableRole"] === ROW_ROLE;
const isCell = (node: PMNode): boolean => CELL_ROLES.has(String(node.type.spec["tableRole"]));
const hasHiddenRow = (table: PMNode): boolean => {
  for (let index = 0; index < table.childCount; index++) {
    if (table.child(index).attrs["hidden"] === true) return true;
  }
  return false;
};

const hasPendingRevision = (root: PMNode): boolean => {
  let pending = false;
  const inspect = (node: PMNode): boolean => {
    if (
      node.attrs["tblPrChange"] != null ||
      node.attrs["trIns"] != null ||
      node.attrs["trDel"] != null ||
      node.attrs["trPrChange"] != null ||
      node.attrs["cellMarker"] != null ||
      node.attrs["tcPrChange"] != null ||
      node.attrs["pPrMark"] != null ||
      node.attrs["_propertyChanges"] != null ||
      node.attrs["_suggestedInsert"] != null ||
      node.marks.some(
        ({ type }) =>
          type.name === "insertion" ||
          type.name === "deletion" ||
          type.name === "runPropertyChange",
      )
    ) {
      pending = true;
      return false;
    }
    return true;
  };
  if (!inspect(root)) return true;
  root.descendants(inspect);
  return pending;
};

const structuralState = (table: PMNode): string => {
  const rows: unknown[] = [];
  table.forEach((row) => {
    const cells: unknown[] = [];
    row.forEach((cell) =>
      cells.push({
        role: String(cell.type.spec["tableRole"]),
        colspan: cell.attrs["colspan"] ?? 1,
        rowspan: cell.attrs["rowspan"] ?? 1,
      }),
    );
    rows.push({ role: String(row.type.spec["tableRole"]), cells });
  });
  return canonicalJson({
    role: String(table.type.spec["tableRole"]),
    columnWidths: table.attrs["columnWidths"] ?? null,
    rows,
  });
};

const simpleRectangularWidth = (table: PMNode): number | null => {
  const map = TableMap.get(table);
  if (map.problems !== null || map.width < 1 || map.height !== table.childCount) return null;
  for (let rowIndex = 0; rowIndex < table.childCount; rowIndex++) {
    const row = table.child(rowIndex);
    if (!isRow(row) || row.childCount !== map.width) return null;
    for (let cellIndex = 0; cellIndex < row.childCount; cellIndex++) {
      const cell = row.child(cellIndex);
      if (!isCell(cell) || cell.attrs["colspan"] !== 1 || cell.attrs["rowspan"] !== 1) {
        return null;
      }
    }
  }
  return map.width;
};

const sourceTableIndexOf = (
  payload: ResolvedDocxTableStructureOperandPayload,
  sourceSnapshot: ResolvedDocxStorySnapshot,
): number | null => {
  const source = (() => {
    switch (payload.type) {
      case "insertTableRow":
      case "insertTableColumn":
        return payload.anchor.source;
      case "deleteTable":
      case "deleteTableRow":
      case "deleteTableColumn":
      case "replaceTable":
        return payload.source;
      case "insertTable":
        return null;
      default: {
        const unreachable: never = payload;
        return panic("Unhandled table operand while resolving source ownership", {
          payload: unreachable,
        });
      }
    }
  })();
  if (!source) return null;
  return (
    resolvedDocxSourceOperandBlock(source, sourceSnapshot).table?.tableIndex ??
    panic("A table-structure source block has no exact table ownership")
  );
};

const targetTableIndexOf = (payload: ResolvedDocxTableStructureOperandPayload): number | null => {
  switch (payload.type) {
    case "insertTable":
    case "insertTableRow":
    case "insertTableColumn":
      return payload.change.tableIndex;
    case "replaceTable":
      return payload.owner.type === "canonical-replacement"
        ? payload.owner.replacement.revisedTableIndex
        : payload.owner.inserted.tableIndex;
    case "deleteTable":
    case "deleteTableRow":
    case "deleteTableColumn":
      return null;
    default: {
      const unreachable: never = payload;
      return panic("Unhandled table operand while resolving target ownership", {
        payload: unreachable,
      });
    }
  }
};

const targetTableIndexForBase = (
  pairings: readonly TableGeometryPairing[],
  baseTableIndex: number,
): number | null => {
  let target: number | null = null;
  for (const pairing of pairings) {
    if (pairing.base.tableIndex !== baseTableIndex) continue;
    if (target !== null && target !== pairing.target.tableIndex) return null;
    target = pairing.target.tableIndex;
  }
  return target;
};

const tablePairings = (
  pairings: readonly TableGeometryPairing[],
  baseTableIndex: number,
  targetTableIndex: number,
): readonly TableGeometryPairing[] =>
  pairings.filter(
    ({ base, target }) =>
      base.tableIndex === baseTableIndex && target.tableIndex === targetTableIndex,
  );

const ownSourceTable = (
  tableIndex: number,
  tables: ReadonlyMap<number, { readonly start: number; readonly node: PMNode }>,
  doc: PMNode,
): SourceTable | TableStructurePreflightResult => {
  const descriptor = tables.get(tableIndex);
  if (!descriptor) return unsupported({ reason: "missing-source-table" });
  const live = doc.nodeAt(descriptor.start);
  if (!live || !isTable(live)) return unsupported({ reason: "missing-source-table" });
  if (live !== descriptor.node && !live.eq(descriptor.node)) {
    return unsupported({ reason: "source-table-mismatch" });
  }
  if (hasPendingRevision(live)) {
    return unsupported({ reason: "pending-structural-revision", side: "source" });
  }
  if (hasHiddenRow(live)) {
    return unsupported({ reason: "unprojected-table-structure", side: "source" });
  }
  return Object.freeze({
    index: tableIndex,
    position: descriptor.start,
    expected: live,
    structuralState: structuralState(live),
  });
};

const ownTargetTable = (
  targetSnapshot: ResolvedDocxStorySnapshot,
  tableIndex: number,
): PMNode | TableStructurePreflightResult => {
  const target = findResolvedDocxTableNode(targetSnapshot, tableIndex);
  if (!target) return unsupported({ reason: "missing-target-table" });
  if (!isTable(target)) return unsupported({ reason: "unexpected-node-role", side: "target" });
  if (hasPendingRevision(target)) {
    return unsupported({ reason: "pending-structural-revision", side: "target" });
  }
  if (hasHiddenRow(target)) {
    return unsupported({ reason: "unprojected-table-structure", side: "target" });
  }
  return target;
};

const portableTarget = (target: PMNode): TableStructurePreflightResult | PMNode => {
  if (!tableTemplateCanCrossPackageLosslessly(target)) {
    return unsupported({ reason: "package-bound-template", side: "target" });
  }
  return target;
};

const insertionBoundary = (
  source: ResolvedDocxSourceOperand,
  position: "after" | "before",
  sourceSnapshot: ResolvedDocxStorySnapshot,
  resolver: FolioStableBlockResolver,
  target: PMNode,
  doc: PMNode,
): InsertionBoundary | TableStructurePreflightResult => {
  const block = resolvedDocxSourceOperandBlock(source, sourceSnapshot);
  const resolved = resolver.resolve(block.identity.id);
  if (resolved.type === "unsupported") {
    return unsupported({ reason: "invalid-insertion-boundary" });
  }
  const at = position === "before" ? resolved.blockFrom : resolved.blockTo;
  const boundary = doc.resolve(at);
  const index = boundary.index();
  if (
    boundary.posAtIndex(index) !== at ||
    !boundary.parent.canReplaceWith(index, index, target.type, target.marks)
  ) {
    return unsupported({ reason: "invalid-insertion-boundary" });
  }
  return Object.freeze({
    position: at,
    depth: boundary.depth,
    parentTypeName: boundary.parent.type.name,
  });
};

const terminalCarrier = (
  source: ResolvedDocxSourceOperand | undefined,
  sourceTable: SourceTable | null,
  sourceSnapshot: ResolvedDocxStorySnapshot,
  resolver: FolioStableBlockResolver,
  doc: PMNode,
): TerminalCarrier | TableStructurePreflightResult | undefined => {
  if (!source) return undefined;
  const block = resolvedDocxSourceOperandBlock(source, sourceSnapshot);
  const resolved = resolver.resolve(block.identity.id);
  if (resolved.type === "unsupported" || resolved.blockNode.type.name !== "paragraph") {
    return unsupported({ reason: "terminal-carrier-missing" });
  }
  if (resolved.blockNode.content.size !== 0 || hasPendingRevision(resolved.blockNode)) {
    return unsupported({ reason: "terminal-carrier-not-empty" });
  }
  if (resolved.blockTo !== doc.content.size) {
    return unsupported({ reason: "terminal-carrier-not-final" });
  }
  if (doc.resolve(resolved.blockFrom).depth !== 0) {
    return unsupported({ reason: "terminal-source-not-body-peer" });
  }
  if (sourceTable && sourceTable.position + sourceTable.expected.nodeSize !== resolved.blockFrom) {
    return unsupported({ reason: "terminal-carrier-missing" });
  }
  return Object.freeze({ position: resolved.blockFrom, expected: resolved.blockNode });
};

const rowMapping = (
  pairings: readonly TableGeometryPairing[],
): ReadonlyMap<number, number> | null => {
  const mapping = new Map<number, number>();
  for (const { base, target } of pairings) {
    const existing = mapping.get(base.rowIndex);
    if (existing !== undefined && existing !== target.rowIndex) return null;
    mapping.set(base.rowIndex, target.rowIndex);
  }
  return mapping;
};

type GridCell = {
  readonly node: PMNode;
  readonly left: number;
  readonly right: number;
  readonly top: number;
  readonly bottom: number;
};

type TableGrid = {
  readonly map: TableMap;
  readonly cellsByPhysicalCoordinate: ReadonlyMap<string, GridCell>;
  readonly cellsByOffset: ReadonlyMap<number, GridCell>;
};

const tableGrid = (table: PMNode): TableGrid | null => {
  if (!isTable(table)) return null;
  const map = TableMap.get(table);
  if (map.problems !== null || map.width < 1 || map.height !== table.childCount) return null;
  const rectanglesByOffset = new Map<
    number,
    { left: number; right: number; top: number; bottom: number }
  >();
  for (const [index, offset] of map.map.entries()) {
    const column = index % map.width;
    const row = Math.floor(index / map.width);
    const rectangle = rectanglesByOffset.get(offset);
    if (rectangle) {
      rectangle.left = Math.min(rectangle.left, column);
      rectangle.right = Math.max(rectangle.right, column + 1);
      rectangle.top = Math.min(rectangle.top, row);
      rectangle.bottom = Math.max(rectangle.bottom, row + 1);
      continue;
    }
    rectanglesByOffset.set(offset, {
      left: column,
      right: column + 1,
      top: row,
      bottom: row + 1,
    });
  }
  const cellsByPhysicalCoordinate = new Map<string, GridCell>();
  const cellsByOffset = new Map<number, GridCell>();
  let rowOffset = 0;
  for (let rowIndex = 0; rowIndex < table.childCount; rowIndex++) {
    const row = table.child(rowIndex);
    if (!isRow(row)) return null;
    let cellOffset = rowOffset + 1;
    for (let cellIndex = 0; cellIndex < row.childCount; cellIndex++) {
      const node = row.child(cellIndex);
      if (!isCell(node)) return null;
      const rectangle = rectanglesByOffset.get(cellOffset);
      if (!rectangle) return null;
      const cell = Object.freeze({
        node,
        left: rectangle.left,
        right: rectangle.right,
        top: rectangle.top,
        bottom: rectangle.bottom,
      });
      cellsByPhysicalCoordinate.set(`${String(rowIndex)}:${String(cellIndex)}`, cell);
      cellsByOffset.set(cellOffset, cell);
      cellOffset += node.nodeSize;
    }
    rowOffset += row.nodeSize;
  }
  if (cellsByOffset.size !== rectanglesByOffset.size) return null;
  return Object.freeze({ map, cellsByPhysicalCoordinate, cellsByOffset });
};

const gridCellAt = (
  grid: TableGrid,
  rowIndex: number,
  columnIndex: number,
): GridCell | null => {
  if (
    rowIndex < 0 ||
    rowIndex >= grid.map.height ||
    columnIndex < 0 ||
    columnIndex >= grid.map.width
  ) {
    return null;
  }
  const offset = grid.map.map[rowIndex * grid.map.width + columnIndex];
  if (offset === undefined) return null;
  return grid.cellsByOffset.get(offset) ?? null;
};

const insertedColumnCells = (
  grid: TableGrid,
  columnIndex: number,
): readonly PMNode[] | null => {
  const cells: PMNode[] = [];
  for (let rowIndex = 0; rowIndex < grid.map.height; rowIndex++) {
    const cell = gridCellAt(grid, rowIndex, columnIndex);
    if (
      !cell ||
      cell.left !== columnIndex ||
      cell.right !== columnIndex + 1 ||
      cell.top !== rowIndex ||
      cell.bottom !== rowIndex + 1
    ) {
      return null;
    }
    cells.push(cell.node);
  }
  return Object.freeze(cells);
};

const boundaryCutsCell = (grid: TableGrid, boundary: number): boolean => {
  if (boundary <= 0 || boundary >= grid.map.width) return false;
  for (let rowIndex = 0; rowIndex < grid.map.height; rowIndex++) {
    const left = grid.map.map[rowIndex * grid.map.width + boundary - 1];
    const right = grid.map.map[rowIndex * grid.map.width + boundary];
    if (left === right) return true;
  }
  return false;
};

const columnMapping = (
  pairings: readonly TableGeometryPairing[],
  sourceGrid: TableGrid,
  targetGrid: TableGrid,
): ReadonlyMap<number, number> | null => {
  const mapping = new Map<number, number>();
  for (const { base, target } of pairings) {
    const baseCell = sourceGrid.cellsByPhysicalCoordinate.get(
      `${String(base.rowIndex)}:${String(base.cellIndex)}`,
    );
    const targetCell = targetGrid.cellsByPhysicalCoordinate.get(
      `${String(target.rowIndex)}:${String(target.cellIndex)}`,
    );
    if (
      !baseCell ||
      !targetCell ||
      baseCell.top !== targetCell.top ||
      baseCell.bottom !== targetCell.bottom ||
      baseCell.right - baseCell.left !== targetCell.right - targetCell.left
    ) {
      return null;
    }
    for (let offset = 0; offset < baseCell.right - baseCell.left; offset++) {
      const sourceColumn = baseCell.left + offset;
      const targetColumn = targetCell.left + offset;
      const existing = mapping.get(sourceColumn);
      if (existing !== undefined && existing !== targetColumn) return null;
      mapping.set(sourceColumn, targetColumn);
    }
  }
  return mapping;
};

const validateRowEdits = (
  source: PMNode,
  target: PMNode,
  edits: readonly RowEdit[],
  pairings: readonly TableGeometryPairing[],
): TableStructureUnsupportedIssue | null => {
  if (simpleRectangularWidth(source) === null) {
    return { reason: "unrepresentable-span", side: "source" };
  }
  if (simpleRectangularWidth(target) === null) {
    return { reason: "unrepresentable-span", side: "target" };
  }
  const deleted = new Set<number>();
  const inserted = new Set<number>();
  const insertionsByBoundary = new Map<number, Extract<RowEdit, { readonly type: "insert" }>[]>();
  for (const edit of edits) {
    if (edit.type === "delete") {
      if (edit.baseRowIndex < 0 || edit.baseRowIndex >= source.childCount) {
        return { reason: "invalid-structural-coordinate" };
      }
      if (deleted.has(edit.baseRowIndex)) return { reason: "duplicate-structural-coordinate" };
      deleted.add(edit.baseRowIndex);
      continue;
    }
    if (
      edit.baseBoundaryIndex < 0 ||
      edit.baseBoundaryIndex > source.childCount ||
      edit.targetRowIndex < 0 ||
      edit.targetRowIndex >= target.childCount
    ) {
      return { reason: "invalid-structural-coordinate" };
    }
    if (inserted.has(edit.targetRowIndex)) return { reason: "duplicate-structural-coordinate" };
    inserted.add(edit.targetRowIndex);
    const at = insertionsByBoundary.get(edit.baseBoundaryIndex) ?? [];
    at.push(edit);
    insertionsByBoundary.set(edit.baseBoundaryIndex, at);
  }
  const mapping = rowMapping(pairings);
  if (!mapping) return { reason: "non-reconstructable-structure" };
  const projected: (
    | { readonly type: "base"; readonly index: number }
    | { readonly type: "target"; readonly index: number }
  )[] = [];
  for (let boundary = 0; boundary <= source.childCount; boundary++) {
    for (const edit of (insertionsByBoundary.get(boundary) ?? []).toSorted(
      (left, right) => left.targetRowIndex - right.targetRowIndex,
    )) {
      projected.push({ type: "target", index: edit.targetRowIndex });
    }
    if (boundary < source.childCount && !deleted.has(boundary)) {
      projected.push({ type: "base", index: boundary });
    }
  }
  if (projected.length !== target.childCount) return { reason: "non-reconstructable-structure" };
  for (const [targetIndex, entry] of projected.entries()) {
    if (entry.type === "target") {
      if (entry.index !== targetIndex) {
        return { reason: "non-reconstructable-structure" };
      }
      continue;
    }
    if (mapping.get(entry.index) !== targetIndex) {
      return { reason: "non-reconstructable-structure" };
    }
  }
  return null;
};

type ValidateColumnEditsOptions = {
  readonly source: PMNode;
  readonly target: PMNode;
  readonly sourceGrid: TableGrid;
  readonly targetGrid: TableGrid;
  readonly edits: readonly ColumnEdit[];
  readonly pairings: readonly TableGeometryPairing[];
  readonly tableGridTransition?: TableGeometryPairing;
};

type ValidatedColumnEdits =
  | {
      readonly status: "ready";
      readonly trackedColumnWidths?: readonly number[];
    }
  | { readonly status: "unsupported"; readonly issue: TableStructureUnsupportedIssue };

const invalidColumnEdits = (issue: TableStructureUnsupportedIssue): ValidatedColumnEdits => ({
  status: "unsupported",
  issue,
});

const validateColumnEdits = ({
  source,
  target,
  sourceGrid,
  targetGrid,
  edits,
  pairings,
  tableGridTransition,
}: ValidateColumnEditsOptions): ValidatedColumnEdits => {
  const sourceWidth = sourceGrid.map.width;
  const targetWidth = targetGrid.map.width;
  if (source.childCount !== target.childCount) {
    return invalidColumnEdits({ reason: "non-reconstructable-structure" });
  }
  const deleted = new Set<number>();
  const inserted = new Set<number>();
  const insertionsByBoundary = new Map<
    number,
    Extract<ColumnEdit, { readonly type: "insert" }>[]
  >();
  for (const edit of edits) {
    if (edit.type === "delete") {
      if (edit.baseColumnIndex < 0 || edit.baseColumnIndex >= sourceWidth) {
        return invalidColumnEdits({ reason: "invalid-structural-coordinate" });
      }
      for (let rowIndex = 0; rowIndex < sourceGrid.map.height; rowIndex++) {
        const cell = gridCellAt(sourceGrid, rowIndex, edit.baseColumnIndex);
        if (
          !cell ||
          cell.left !== edit.baseColumnIndex ||
          cell.right !== edit.baseColumnIndex + 1
        ) {
          return invalidColumnEdits({ reason: "unrepresentable-span", side: "source" });
        }
      }
      if (deleted.has(edit.baseColumnIndex)) {
        return invalidColumnEdits({ reason: "duplicate-structural-coordinate" });
      }
      deleted.add(edit.baseColumnIndex);
      continue;
    }
    if (
      edit.baseBoundaryIndex < 0 ||
      edit.baseBoundaryIndex > sourceWidth ||
      edit.targetColumnIndex < 0 ||
      edit.targetColumnIndex >= targetWidth
    ) {
      return invalidColumnEdits({ reason: "invalid-structural-coordinate" });
    }
    if (boundaryCutsCell(sourceGrid, edit.baseBoundaryIndex)) {
      return invalidColumnEdits({ reason: "unrepresentable-span", side: "source" });
    }
    if (inserted.has(edit.targetColumnIndex)) {
      return invalidColumnEdits({ reason: "duplicate-structural-coordinate" });
    }
    inserted.add(edit.targetColumnIndex);
    const at = insertionsByBoundary.get(edit.baseBoundaryIndex) ?? [];
    at.push(edit);
    insertionsByBoundary.set(edit.baseBoundaryIndex, at);
  }
  const mapping = columnMapping(pairings, sourceGrid, targetGrid);
  if (!mapping) return invalidColumnEdits({ reason: "non-reconstructable-structure" });
  const projected: (
    | { readonly type: "base"; readonly index: number }
    | { readonly type: "target"; readonly index: number }
  )[] = [];
  for (let boundary = 0; boundary <= sourceWidth; boundary++) {
    for (const edit of (insertionsByBoundary.get(boundary) ?? []).toSorted(
      (left, right) => left.targetColumnIndex - right.targetColumnIndex,
    )) {
      projected.push({ type: "target", index: edit.targetColumnIndex });
    }
    if (boundary < sourceWidth && !deleted.has(boundary)) {
      projected.push({ type: "base", index: boundary });
    }
  }
  if (projected.length !== targetWidth) {
    return invalidColumnEdits({ reason: "non-reconstructable-structure" });
  }
  for (const [targetIndex, entry] of projected.entries()) {
    if (
      (entry.type === "target" && entry.index !== targetIndex) ||
      (entry.type === "base" && mapping.get(entry.index) !== targetIndex)
    ) {
      return invalidColumnEdits({ reason: "non-reconstructable-structure" });
    }
  }

  const sourceColumnWidths = expectTableAttrs(source).columnWidths ?? null;
  const targetColumnWidths = expectTableAttrs(target).columnWidths ?? null;
  if (sourceColumnWidths === null && targetColumnWidths === null) {
    if (tableGridTransition) {
      return panic("A delegated table-grid transition no longer names a grid difference");
    }
    return { status: "ready" };
  }
  const invalidGridTransition = (): ValidatedColumnEdits =>
    invalidColumnEdits(
      tableGridTransition
        ? {
            reason: "unrepresentable-table-geometry",
            issue: {
              reason: "non-reconstructable-structure-change",
              scope: "table",
              property: "column-widths",
              base: tableGridTransition.base,
              target: tableGridTransition.target,
            },
          }
        : { reason: "non-reconstructable-structure" },
    );
  if (
    sourceColumnWidths === null ||
    targetColumnWidths === null ||
    sourceColumnWidths.length !== sourceWidth ||
    targetColumnWidths.length !== targetWidth
  ) {
    return invalidGridTransition();
  }
  for (const [targetIndex, entry] of projected.entries()) {
    if (
      entry.type === "base" &&
      sourceColumnWidths[entry.index] !== targetColumnWidths[targetIndex]
    ) {
      return invalidGridTransition();
    }
  }
  if (inserted.size === 0) return { status: "ready" };

  const trackedColumnWidths: number[] = [];
  for (let boundary = 0; boundary <= sourceWidth; boundary++) {
    for (const edit of (insertionsByBoundary.get(boundary) ?? []).toSorted(
      (left, right) => left.targetColumnIndex - right.targetColumnIndex,
    )) {
      const width = targetColumnWidths[edit.targetColumnIndex];
      if (width === undefined) {
        return invalidGridTransition();
      }
      trackedColumnWidths.push(width);
    }
    if (boundary < sourceWidth) {
      const width = sourceColumnWidths[boundary];
      if (width === undefined) {
        return invalidGridTransition();
      }
      trackedColumnWidths.push(width);
    }
  }
  return { status: "ready", trackedColumnWidths: Object.freeze(trackedColumnWidths) };
};

type MutableTableGroup = {
  readonly source: SourceTable;
  targetTableIndex: number | null;
  axis: "row" | "column" | null;
  whole: OwnedWholeTableInstruction | null;
  readonly rowEdits: RowEdit[];
  readonly columnEdits: ColumnEdit[];
};

const taskFor = (instruction: OwnedInstruction): PreparedTableStructureTask => {
  const operands = (() => {
    switch (instruction.type) {
      case "insertTable":
      case "deleteTable":
      case "replaceTable":
        return [instruction.operand];
      case "editTableRows":
      case "editTableColumns":
        return instruction.edits.map(({ operand }) => operand);
      default: {
        const unreachable: never = instruction;
        return panic("Unhandled owned table instruction", { instruction: unreachable });
      }
    }
  })();
  let schedule: PreparedTableStructureTask["schedule"];
  if (instruction.type === "insertTable") {
    schedule = Object.freeze({ phase: "insertion", position: instruction.boundary.position });
  } else {
    let to = instruction.source.position + instruction.source.expected.nodeSize;
    if (instruction.type === "replaceTable" && instruction.terminalCarrier) {
      to = instruction.terminalCarrier.position + instruction.terminalCarrier.expected.nodeSize;
    }
    schedule = Object.freeze({
      phase: "source",
      position: instruction.source.position,
      from: instruction.source.position,
      to,
    });
  }
  return Object.freeze({
    [PREPARED_TABLE_STRUCTURE_TASK]: true as const,
    schedule,
    operands: Object.freeze(operands),
  });
};

const validLimits = (
  limits: TableStructurePreflightLimits,
): Extract<TableStructurePreflightResult, { readonly status: "unsupported" }> | null => {
  for (const limit of [
    "maxOperands",
    "maxStructuralEdits",
    "maxDocumentUnits",
    "maxTargetTemplateUnits",
    "maxTables",
  ] as const satisfies readonly TableStructureLimit[]) {
    if (!Number.isSafeInteger(limits[limit]) || limits[limit] < 0) {
      return unsupported({ reason: "invalid-limit", limit, actual: limits[limit] });
    }
  }
  return null;
};

type TableStructurePreflightContext = {
  readonly doc: PMNode;
  readonly comparison: ResolvedDocxStoryComparison;
  readonly limits: TableStructurePreflightLimits;
  readonly baseSnapshot: ResolvedDocxStorySnapshot;
  readonly targetSnapshot: ResolvedDocxStorySnapshot;
  readonly baseTables: ReadonlyMap<number, { readonly start: number; readonly node: PMNode }>;
  readonly resolver: FolioStableBlockResolver;
  readonly pairings: readonly TableGeometryPairing[];
  readonly countedTargets: Set<PMNode>;
  readonly sourceCache: Map<number, SourceTable>;
  readonly tableGridCache: Map<PMNode, TableGrid | null>;
  structuralEdits: number;
  targetTemplateUnits: number;
};

type TableStructurePreflightContextResult =
  | { readonly status: "ready"; readonly context: TableStructurePreflightContext }
  | { readonly status: "unsupported"; readonly issue: TableStructureUnsupportedIssue };

const createTableStructurePreflightContext = ({
  doc,
  comparison,
  operands,
  limits,
}: {
  readonly doc: PMNode;
  readonly comparison: ResolvedDocxStoryComparison;
  readonly operands: readonly ResolvedDocxTableOperand[];
  readonly limits: TableStructurePreflightLimits;
}): TableStructurePreflightContextResult => {
  const invalidLimits = validLimits(limits);
  if (invalidLimits) return invalidLimits;
  if (operands.length > limits.maxOperands) {
    return unsupported({
      reason: "limit-exceeded",
      limit: "maxOperands",
      maximum: limits.maxOperands,
      actual: operands.length,
    });
  }
  if (doc.nodeSize > limits.maxDocumentUnits) {
    return unsupported({
      reason: "limit-exceeded",
      limit: "maxDocumentUnits",
      maximum: limits.maxDocumentUnits,
      actual: doc.nodeSize,
    });
  }
  if (new Set(operands).size !== operands.length) {
    return unsupported({ reason: "duplicate-operand" });
  }
  const structuralEdits = operands.filter(({ type }) => type !== "matchTableFormatting").length;
  if (structuralEdits > limits.maxStructuralEdits) {
    return unsupported({
      reason: "limit-exceeded",
      limit: "maxStructuralEdits",
      maximum: limits.maxStructuralEdits,
      actual: structuralEdits,
    });
  }

  const { baseSnapshot, targetSnapshot } = resolvedDocxStoryComparisonPayload(comparison);
  const operationSnapshot = resolvedDocxOperationSnapshot(baseSnapshot);
  const storyTables = storyTablesOf(operationSnapshot);
  if (storyTables.length > limits.maxTables) {
    return unsupported({
      reason: "limit-exceeded",
      limit: "maxTables",
      maximum: limits.maxTables,
      actual: storyTables.length,
    });
  }
  return {
    status: "ready",
    context: {
      doc,
      comparison,
      limits,
      baseSnapshot,
      targetSnapshot,
      baseTables: new Map(storyTables.map(({ index, start, node }) => [index, { start, node }])),
      resolver: FolioStableBlockResolver.create(doc, operationSnapshot),
      pairings: resolvedDocxTableGeometryPairings(comparison),
      countedTargets: new Set(),
      sourceCache: new Map(),
      tableGridCache: new Map(),
      structuralEdits: 0,
      targetTemplateUnits: 0,
    },
  };
};

/**
 * Resolve every table operation through its exact comparison capsule before a
 * caller-owned transaction exists. The resulting tasks retain only nominal
 * operands; raw coordinates and target nodes stay closure-owned here.
 */
const preflightTableStructureComponent = ({
  context,
  operands,
}: {
  readonly context: TableStructurePreflightContext;
  readonly operands: readonly ResolvedDocxTableOperand[];
}): TableStructurePreflightResult => {
  const { doc, comparison, limits, baseSnapshot, targetSnapshot, baseTables, resolver, pairings } =
    context;
  const formatOperands = operands.filter(
    (operand): operand is ResolvedDocxTableFormatOperand => operand.type === "matchTableFormatting",
  );
  if (formatOperands.length > 1) return unsupported({ reason: "duplicate-geometry-operand" });
  const structuralOperands = operands.filter(
    (operand): operand is ResolvedDocxTableStructureOperand =>
      operand.type !== "matchTableFormatting",
  );
  const structuralOperations = structuralOperands.map((operand) =>
    resolveTableStructureOperation(operand, comparison),
  );
  const geometryOperand = formatOperands.at(0);
  const geometryPayload = geometryOperand
    ? resolvedDocxTableFormatOperandPayload(geometryOperand, comparison)
    : null;
  if (geometryPayload?.status === "unsupported") return unsupported(geometryPayload.issue);
  const pendingTableGridTransitions = new Map<string, TableGeometryPairing>();
  if (geometryPayload?.status === "ready") {
    for (const transition of tableGeometryProgramTableGridTransitions(geometryPayload.program)) {
      const key = `${String(transition.base.tableIndex)}:${String(transition.target.tableIndex)}`;
      if (pendingTableGridTransitions.has(key)) {
        return panic("A table geometry program delegated the same grid transition twice");
      }
      pendingTableGridTransitions.set(key, transition);
    }
  }

  const countTarget = (target: PMNode): TableStructurePreflightResult | null => {
    if (context.countedTargets.has(target)) return null;
    context.countedTargets.add(target);
    context.targetTemplateUnits += target.nodeSize;
    if (
      !Number.isSafeInteger(context.targetTemplateUnits) ||
      context.targetTemplateUnits > limits.maxTargetTemplateUnits
    ) {
      return unsupported({
        reason: "limit-exceeded",
        limit: "maxTargetTemplateUnits",
        maximum: limits.maxTargetTemplateUnits,
        actual: context.targetTemplateUnits,
      });
    }
    return null;
  };

  const sourceTable = (index: number): SourceTable | TableStructurePreflightResult => {
    const cached = context.sourceCache.get(index);
    if (cached) return cached;
    const owned = ownSourceTable(index, baseTables, doc);
    if ("status" in owned) return owned;
    context.sourceCache.set(index, owned);
    return owned;
  };

  const gridFor = (table: PMNode): TableGrid | null => {
    if (context.tableGridCache.has(table)) return context.tableGridCache.get(table) ?? null;
    const grid = tableGrid(table);
    context.tableGridCache.set(table, grid);
    return grid;
  };

  const groups = new Map<number, MutableTableGroup>();
  const insertions: OwnedTableInsertion[] = [];
  const groupFor = (source: SourceTable): MutableTableGroup => {
    const existing = groups.get(source.index);
    if (existing) return existing;
    const created: MutableTableGroup = {
      source,
      targetTableIndex: null,
      axis: null,
      whole: null,
      rowEdits: [],
      columnEdits: [],
    };
    groups.set(source.index, created);
    return created;
  };

  for (const operation of structuralOperations) {
    context.structuralEdits++;
    if (context.structuralEdits > limits.maxStructuralEdits) {
      return unsupported({
        reason: "limit-exceeded",
        limit: "maxStructuralEdits",
        maximum: limits.maxStructuralEdits,
        actual: context.structuralEdits,
      });
    }
    if (operation.type === "insertTable") {
      const { operand, payload } = operation;
      const targetResult = ownTargetTable(targetSnapshot, payload.change.tableIndex);
      if ("status" in targetResult) return targetResult;
      const target = portableTarget(targetResult);
      if ("status" in target) return target;
      const counted = countTarget(target);
      if (counted) return counted;
      const boundary = insertionBoundary(
        payload.anchor.source,
        payload.anchor.position,
        baseSnapshot,
        resolver,
        target,
        doc,
      );
      if ("status" in boundary) return boundary;
      const carrier = terminalCarrier(
        payload.terminalCarrier?.source,
        null,
        baseSnapshot,
        resolver,
        doc,
      );
      if (carrier && "status" in carrier) return carrier;
      if (carrier && boundary.position !== carrier.position + carrier.expected.nodeSize) {
        return unsupported({ reason: "terminal-carrier-missing" });
      }
      insertions.push({
        type: "insertTable",
        operand,
        boundary,
        target,
        ...(carrier && { terminalCarrier: carrier }),
      });
      continue;
    }

    const baseIndex = sourceTableIndexOf(operation.payload, baseSnapshot);
    if (baseIndex === null) return unsupported({ reason: "missing-source-table" });
    const source = sourceTable(baseIndex);
    if ("status" in source) return source;
    const group = groupFor(source);
    if (operation.type === "deleteTable") {
      if (group.whole || group.rowEdits.length > 0 || group.columnEdits.length > 0) {
        return unsupported({ reason: "duplicate-source-table" });
      }
      group.whole = { type: "deleteTable", operand: operation.operand, source };
      continue;
    }
    if (operation.type === "replaceTable") {
      const { operand, payload } = operation;
      if (group.whole || group.rowEdits.length > 0 || group.columnEdits.length > 0) {
        return unsupported({ reason: "duplicate-source-table" });
      }
      const revisedTableIndex =
        payload.owner.type === "canonical-replacement"
          ? payload.owner.replacement.revisedTableIndex
          : payload.owner.inserted.tableIndex;
      const targetResult = ownTargetTable(targetSnapshot, revisedTableIndex);
      if ("status" in targetResult) return targetResult;
      const target = portableTarget(targetResult);
      if ("status" in target) return target;
      const counted = countTarget(target);
      if (counted) return counted;
      const carrier = terminalCarrier(
        payload.terminalCarrier?.source,
        source,
        baseSnapshot,
        resolver,
        doc,
      );
      if (carrier && "status" in carrier) return carrier;
      group.whole = {
        type: "replaceTable",
        operand,
        source,
        target,
        ...(carrier && { terminalCarrier: carrier }),
      };
      continue;
    }

    const targetIndex =
      targetTableIndexOf(operation.payload) ?? targetTableIndexForBase(pairings, source.index);
    if (targetIndex === null) return unsupported({ reason: "non-reconstructable-structure" });
    if (group.targetTableIndex !== null && group.targetTableIndex !== targetIndex) {
      return unsupported({ reason: "non-reconstructable-structure" });
    }
    group.targetTableIndex = targetIndex;
    const target = ownTargetTable(targetSnapshot, targetIndex);
    if ("status" in target) return target;
    const counted = countTarget(target);
    if (counted) return counted;
    if (operation.type === "insertTableRow" || operation.type === "deleteTableRow") {
      if (group.axis === "column") return unsupported({ reason: "mixed-table-structure-axes" });
      group.axis = "row";
      if (operation.type === "insertTableRow") {
        const { operand, payload } = operation;
        const anchor = resolvedDocxSourceOperandBlock(payload.anchor.source, baseSnapshot).table;
        if (!anchor) return unsupported({ reason: "missing-source-table" });
        const targetRow = target.child(payload.change.rowIndex);
        if (!targetRow || !isRow(targetRow)) {
          return unsupported({ reason: "invalid-structural-coordinate" });
        }
        if (!tableTemplateCanCrossPackageLosslessly(targetRow)) {
          return unsupported({ reason: "package-bound-template", side: "target" });
        }
        group.rowEdits.push({
          type: "insert",
          operand,
          baseBoundaryIndex: anchor.rowIndex + (payload.anchor.position === "after" ? 1 : 0),
          targetRowIndex: payload.change.rowIndex,
          targetRow,
        });
      } else {
        group.rowEdits.push({
          type: "delete",
          operand: operation.operand,
          baseRowIndex: operation.payload.change.rowIndex,
        });
      }
      continue;
    }
    if (group.axis === "row") return unsupported({ reason: "mixed-table-structure-axes" });
    group.axis = "column";
    if (operation.type === "insertTableColumn") {
      const { operand, payload } = operation;
      const anchor = resolvedDocxSourceOperandBlock(payload.anchor.source, baseSnapshot).table;
      if (!anchor) return unsupported({ reason: "missing-source-table" });
      const targetGrid = gridFor(target);
      if (!targetGrid) return unsupported({ reason: "unrepresentable-span", side: "target" });
      const targetCells = insertedColumnCells(targetGrid, payload.change.columnIndex);
      if (!targetCells) {
        return unsupported({ reason: "invalid-structural-coordinate" });
      }
      if (targetCells.some((cell) => !tableTemplateCanCrossPackageLosslessly(cell))) {
        return unsupported({ reason: "package-bound-template", side: "target" });
      }
      group.columnEdits.push({
        type: "insert",
        operand,
        baseBoundaryIndex: anchor.gridColumnIndex + (payload.anchor.position === "after" ? 1 : 0),
        targetColumnIndex: payload.change.columnIndex,
        targetCells,
      });
    } else {
      group.columnEdits.push({
        type: "delete",
        operand: operation.operand,
        baseColumnIndex: operation.payload.change.columnIndex,
      });
    }
  }

  const instructions: OwnedInstruction[] = [...insertions];
  for (const group of groups.values()) {
    if (group.whole) {
      instructions.push(group.whole);
      continue;
    }
    const targetIndex = group.targetTableIndex;
    if (targetIndex === null) return unsupported({ reason: "non-reconstructable-structure" });
    const target = ownTargetTable(targetSnapshot, targetIndex);
    if ("status" in target) return target;
    const ownedPairings = tablePairings(pairings, group.source.index, targetIndex);
    if (group.axis === "row") {
      const problem = validateRowEdits(
        group.source.expected,
        target,
        group.rowEdits,
        ownedPairings,
      );
      if (problem) return unsupported(problem);
      instructions.push({
        type: "editTableRows",
        source: group.source,
        target,
        edits: Object.freeze(group.rowEdits),
      });
      continue;
    }
    if (group.axis === "column") {
      const sourceGrid = gridFor(group.source.expected);
      if (!sourceGrid) return unsupported({ reason: "unrepresentable-span", side: "source" });
      const targetGrid = gridFor(target);
      if (!targetGrid) return unsupported({ reason: "unrepresentable-span", side: "target" });
      const tableGridTransitionKey = `${String(group.source.index)}:${String(targetIndex)}`;
      const tableGridTransition = pendingTableGridTransitions.get(tableGridTransitionKey);
      const validated = validateColumnEdits({
        source: group.source.expected,
        target,
        sourceGrid,
        targetGrid,
        edits: group.columnEdits,
        pairings: ownedPairings,
        ...(tableGridTransition === undefined ? {} : { tableGridTransition }),
      });
      if (validated.status === "unsupported") return unsupported(validated.issue);
      pendingTableGridTransitions.delete(tableGridTransitionKey);
      instructions.push({
        type: "editTableColumns",
        source: group.source,
        target,
        edits: Object.freeze(group.columnEdits),
        ...(validated.trackedColumnWidths === undefined
          ? {}
          : { trackedColumnWidths: validated.trackedColumnWidths }),
      });
    }
  }

  const unclaimedGridTransition = pendingTableGridTransitions.values().next().value;
  if (unclaimedGridTransition) {
    return unsupported({
      reason: "unrepresentable-table-geometry",
      issue: {
        reason: "non-reconstructable-structure-change",
        scope: "table",
        property: "column-widths",
        base: unclaimedGridTransition.base,
        target: unclaimedGridTransition.target,
      },
    });
  }

  const sourceInstructions = instructions
    .filter((instruction) => instruction.type !== "insertTable")
    .toSorted((left, right) => left.source.position - right.source.position);
  for (let index = 1; index < sourceInstructions.length; index++) {
    const previous = sourceInstructions[index - 1];
    const current = sourceInstructions[index];
    if (
      previous &&
      current &&
      current.source.position < previous.source.position + previous.source.expected.nodeSize
    ) {
      return unsupported({ reason: "overlapping-table-obligations" });
    }
  }
  for (const insertion of insertions) {
    for (const source of sourceInstructions) {
      if (
        insertion.boundary.position > source.source.position &&
        insertion.boundary.position < source.source.position + source.source.expected.nodeSize
      ) {
        return unsupported({ reason: "overlapping-table-obligations" });
      }
    }
  }

  const geometry =
    geometryOperand && geometryPayload?.status === "ready"
      ? { operand: geometryOperand, program: geometryPayload.program }
      : null;

  const tasks = instructions.map(taskFor);
  const instructionsByTask = new Map<PreparedTableStructureTask, OwnedInstruction>();
  for (const [index, task] of tasks.entries()) {
    const instruction = instructions[index] ?? panic("A table task lost its owned instruction");
    instructionsByTask.set(task, instruction);
  }
  const program = Object.freeze({
    [PREPARED_TABLE_STRUCTURE_PROGRAM]: true as const,
    tasks: Object.freeze(tasks),
    operands: Object.freeze([...operands]),
    ...(geometryOperand && { geometryOperand }),
  });
  ownedPrograms.set(program, {
    sourceDocument: doc,
    instructionsByTask,
    geometry,
    executedTasks: new Set(),
    geometryExecuted: false,
  });
  return Object.freeze({ status: "ready", program });
};

export const preflightTableStructureProgram = ({
  doc,
  comparison,
  operands,
  limits = DEFAULT_TABLE_STRUCTURE_PREFLIGHT_LIMITS,
}: {
  readonly doc: PMNode;
  readonly comparison: ResolvedDocxStoryComparison;
  readonly operands: readonly ResolvedDocxTableOperand[];
  readonly limits?: TableStructurePreflightLimits;
}): TableStructurePreflightResult => {
  const prepared = createTableStructurePreflightContext({
    doc,
    comparison,
    operands,
    limits,
  });
  return prepared.status === "unsupported"
    ? prepared
    : preflightTableStructureComponent({ context: prepared.context, operands });
};

/** Preflight each canonical table component while sharing all global limits and indexes. */
export const preflightTableStructureComponents = ({
  doc,
  comparison,
  components,
  limits = DEFAULT_TABLE_STRUCTURE_PREFLIGHT_LIMITS,
}: {
  readonly doc: PMNode;
  readonly comparison: ResolvedDocxStoryComparison;
  readonly components: readonly ResolvedDocxTableComponent[];
  readonly limits?: TableStructurePreflightLimits;
}): TableStructureComponentsPreflightResult => {
  if (new Set(components).size !== components.length) {
    return panic("A table component was supplied more than once");
  }
  const componentOperands = components.map((component) => ({
    component,
    operands: resolvedDocxTableComponentOperands(component, comparison),
  }));
  if (componentOperands.some(({ operands }) => operands.length === 0)) {
    return panic("A table component cannot be empty");
  }
  const operands = componentOperands.flatMap(({ operands: ownedOperands }) => ownedOperands);
  const prepared = createTableStructurePreflightContext({
    doc,
    comparison,
    operands,
    limits,
  });
  if (prepared.status === "unsupported") return prepared;

  const results: TableStructureComponentPreflightResult[] = [];
  for (const { component, operands: ownedOperands } of componentOperands) {
    const result = preflightTableStructureComponent({
      context: prepared.context,
      operands: ownedOperands,
    });
    if (result.status === "unsupported" && result.issue.reason === "limit-exceeded") {
      return result;
    }
    results.push(Object.freeze({ component, result }));
  }

  const readyPrograms = results.flatMap(({ result }) =>
    result.status === "ready" ? [result.program] : [],
  );
  const sourceTasks = readyPrograms
    .flatMap(({ tasks }) => tasks)
    .filter(
      (
        task,
      ): task is PreparedTableStructureTask & {
        readonly schedule: Extract<PreparedTableStructureTask["schedule"], { phase: "source" }>;
      } => task.schedule.phase === "source",
    )
    .toSorted((left, right) => left.schedule.from - right.schedule.from);
  for (let index = 1; index < sourceTasks.length; index++) {
    const previous = sourceTasks[index - 1];
    const current = sourceTasks[index];
    if (previous && current && current.schedule.from < previous.schedule.to) {
      return panic("Canonical table components contain overlapping source obligations");
    }
  }
  const insertions = readyPrograms
    .flatMap(({ tasks }) => tasks)
    .filter(({ schedule }) => schedule.phase === "insertion");
  for (const insertion of insertions) {
    for (const source of sourceTasks) {
      if (
        insertion.schedule.position > source.schedule.from &&
        insertion.schedule.position < source.schedule.to
      ) {
        return panic("A canonical table insertion crosses another component's source scope");
      }
    }
  }
  return Object.freeze({ status: "ready", components: Object.freeze(results) });
};

export type TableStructureRevisionStamp = {
  readonly author: string;
  readonly date: string;
};

export type TableStructureAppliedOperand = {
  readonly operand: ResolvedDocxTableOperand;
  readonly revisionIds: readonly number[];
  readonly tableGeometry?: TableGeometryExecutionReceipt;
};

export type TableStructureExecutionResult = {
  readonly transaction: Transaction;
  readonly nextRevisionId: number;
  readonly applied: readonly TableStructureAppliedOperand[];
  readonly localPositionMappingSteps: 0;
};

const ownedProgramOf = (program: PreparedTableStructureProgram): OwnedProgram =>
  ownedPrograms.get(program) ??
  panic("A table-structure program must come from preflightTableStructureProgram");

/** Apply the geometry half only after every structural obligation preflighted. */
export const executeTableStructureGeometry = ({
  tr,
  program,
  revision,
  revisionId,
}: {
  readonly tr: Transaction;
  readonly program: PreparedTableStructureProgram;
  readonly revision: TableStructureRevisionStamp;
  readonly revisionId: number;
}): TableStructureExecutionResult | { readonly issue: TableGeometryExecutionIssue } => {
  const owned = ownedProgramOf(program);
  if (tr.before !== owned.sourceDocument || owned.geometryExecuted) {
    return panic("A table-structure geometry phase was stale or consumed twice");
  }
  owned.geometryExecuted = true;
  if (!owned.geometry) {
    return {
      transaction: tr,
      nextRevisionId: revisionId,
      applied: [],
      localPositionMappingSteps: 0,
    };
  }
  const result = executeTableGeometryProgram({
    tr,
    program: owned.geometry.program,
    revision: { author: revision.author, date: revision.date, idSeed: revisionId },
  });
  if (result.status === "unsupported") return { issue: result.issue };
  return {
    transaction: tr,
    nextRevisionId: result.receipt.nextRevisionId,
    applied: [
      {
        operand: owned.geometry.operand,
        revisionIds: Object.freeze(result.receipt.revisions.map(({ revisionId: id }) => id)),
        tableGeometry: result.receipt,
      },
    ],
    localPositionMappingSteps: 0,
  };
};

const rowPositions = (table: PMNode, tablePosition: number): number[] => {
  const positions: number[] = [];
  let position = tablePosition + 1;
  table.forEach((row) => {
    positions.push(position);
    position += row.nodeSize;
  });
  return positions;
};

const markWholeTable = (
  tr: Transaction,
  table: PMNode,
  tablePosition: number,
  kind: MarkTableRowContentOptions["kind"],
  revision: MarkTableRowContentOptions["revision"],
): void => {
  const attr = kind === "insertion" ? "trIns" : "trDel";
  for (const position of rowPositions(table, tablePosition).toReversed()) {
    tr.setNodeAttribute(position, attr, revision);
    markTableRowContent({ tr, rowPosition: position, kind, revision });
  }
  markStructuralChange(tr);
};

const structuralRevision = (id: number, revision: TableStructureRevisionStamp) => ({
  revisionId: id,
  author: revision.author,
  date: revision.date,
});

const liveSourceTable = (tr: Transaction, source: SourceTable): PMNode => {
  const live = tr.doc.nodeAt(source.position);
  if (!live || !isTable(live) || structuralState(live) !== source.structuralState) {
    return panic("A preflighted table-structure source became stale");
  }
  return live;
};

const setTrackedTableGrid = (
  tr: Transaction,
  tablePosition: number,
  columnWidths: readonly number[],
): void => {
  const table = tr.doc.nodeAt(tablePosition);
  if (!table || !isTable(table) || TableMap.get(table).width !== columnWidths.length) {
    return panic("A preflighted column edit produced an inconsistent tracked table grid");
  }
  const formatting = expectTableAttrs(table)._originalFormatting;
  const originalFormatting = formatting ? { ...formatting } : undefined;
  if (originalFormatting) delete originalFormatting.gridSourceXml;
  tr.setNodeMarkup(
    tablePosition,
    undefined,
    mergeTableAttrs(table, {
      columnWidths: [...columnWidths],
      _originalFormatting: originalFormatting,
    }),
  );
};

/** Execute one closure-owned structural task in the caller's global schedule. */
export const executeTableStructureTask = ({
  tr,
  program,
  task,
  revision,
  revisionId,
}: {
  readonly tr: Transaction;
  readonly program: PreparedTableStructureProgram;
  readonly task: PreparedTableStructureTask;
  readonly revision: TableStructureRevisionStamp;
  readonly revisionId: number;
}): TableStructureExecutionResult => {
  const owned = ownedProgramOf(program);
  if (
    tr.before !== owned.sourceDocument ||
    !owned.geometryExecuted ||
    owned.executedTasks.has(task)
  ) {
    return panic("A table-structure task was stale, out of phase, or consumed twice");
  }
  const instruction =
    owned.instructionsByTask.get(task) ??
    panic("A table-structure task belongs to another program");
  owned.executedTasks.add(task);
  let nextRevisionId = revisionId;
  const applied: TableStructureAppliedOperand[] = [];
  const allocate = (): number => nextRevisionId++;
  switch (instruction.type) {
    case "insertTable": {
      const insertionRevision = structuralRevision(allocate(), revision);
      const liveBoundary = tr.doc.resolve(instruction.boundary.position);
      if (
        liveBoundary.depth !== instruction.boundary.depth ||
        liveBoundary.parent.type.name !== instruction.boundary.parentTypeName
      ) {
        return panic("A preflighted table insertion boundary became stale");
      }
      const table =
        tableFromTemplate({
          schema: tr.doc.type.schema,
          template: instruction.target,
          revision: insertionRevision,
        }) ?? panic("A preflighted table insertion lost its target template");
      tr.insert(instruction.boundary.position, table);
      const revisionIds = [insertionRevision.revisionId];
      if (instruction.terminalCarrier) {
        const carrierRevisionId = allocate();
        tr.setNodeAttribute(instruction.terminalCarrier.position, "pPrMark", {
          kind: "del",
          info: { id: carrierRevisionId, author: revision.author, date: revision.date },
        });
        revisionIds.push(carrierRevisionId);
      }
      markStructuralChange(tr);
      applied.push({ operand: instruction.operand, revisionIds: Object.freeze(revisionIds) });
      break;
    }
    case "deleteTable": {
      const table = liveSourceTable(tr, instruction.source);
      const deletionRevision = structuralRevision(allocate(), revision);
      markWholeTable(tr, table, instruction.source.position, "deletion", deletionRevision);
      applied.push({
        operand: instruction.operand,
        revisionIds: Object.freeze([deletionRevision.revisionId]),
      });
      break;
    }
    case "replaceTable": {
      const table = liveSourceTable(tr, instruction.source);
      const insertionRevision = structuralRevision(allocate(), revision);
      const deletionRevision = structuralRevision(allocate(), revision);
      const target =
        tableFromTemplate({
          schema: tr.doc.type.schema,
          template: instruction.target,
          revision: insertionRevision,
        }) ?? panic("A preflighted table replacement lost its target template");
      markWholeTable(tr, table, instruction.source.position, "deletion", deletionRevision);
      const revisionIds = [insertionRevision.revisionId, deletionRevision.revisionId];
      if (instruction.terminalCarrier) {
        tr.insert(
          instruction.terminalCarrier.position + instruction.terminalCarrier.expected.nodeSize,
          target,
        );
        const carrierRevisionId = allocate();
        tr.setNodeAttribute(instruction.terminalCarrier.position, "pPrMark", {
          kind: "del",
          info: { id: carrierRevisionId, author: revision.author, date: revision.date },
        });
        revisionIds.push(carrierRevisionId);
      } else {
        tr.insert(instruction.source.position, target);
      }
      markStructuralChange(tr);
      applied.push({ operand: instruction.operand, revisionIds: Object.freeze(revisionIds) });
      break;
    }
    case "editTableRows": {
      const table = liveSourceTable(tr, instruction.source);
      const positions = rowPositions(table, instruction.source.position);
      const boundaries = [...positions, instruction.source.position + table.nodeSize - 1];
      for (const edit of instruction.edits) {
        if (edit.type !== "delete") continue;
        const id = allocate();
        const rowPosition =
          positions[edit.baseRowIndex] ?? panic("A preflighted row deletion lost its coordinate");
        const rowRevision = structuralRevision(id, revision);
        tr.setNodeAttribute(rowPosition, "trDel", rowRevision);
        markTableRowContent({
          tr,
          rowPosition,
          kind: "deletion",
          revision: rowRevision,
        });
        applied.push({ operand: edit.operand, revisionIds: Object.freeze([id]) });
      }
      const insertions = instruction.edits
        .filter(
          (edit): edit is Extract<RowEdit, { readonly type: "insert" }> => edit.type === "insert",
        )
        .toSorted(
          (left, right) =>
            right.baseBoundaryIndex - left.baseBoundaryIndex ||
            right.targetRowIndex - left.targetRowIndex,
        );
      for (const edit of insertions) {
        const id = allocate();
        const rowRevision = structuralRevision(id, revision);
        const rowPosition =
          boundaries[edit.baseBoundaryIndex] ??
          panic("A preflighted row insertion lost its boundary");
        const row = edit.targetRow.type.create(
          { ...edit.targetRow.attrs, trIns: rowRevision },
          edit.targetRow.content,
          edit.targetRow.marks,
        );
        tr.insert(rowPosition, row);
        markTableRowContent({ tr, rowPosition, kind: "insertion", revision: rowRevision });
        applied.push({ operand: edit.operand, revisionIds: Object.freeze([id]) });
      }
      markStructuralChange(tr);
      break;
    }
    case "editTableColumns": {
      const table = liveSourceTable(tr, instruction.source);
      const map = TableMap.get(table);
      for (const edit of instruction.edits) {
        if (edit.type !== "delete") continue;
        const id = allocate();
        const columnRevision = structuralRevision(id, revision);
        for (let rowIndex = map.height - 1; rowIndex >= 0; rowIndex--) {
          const cellOffset =
            map.map[rowIndex * map.width + edit.baseColumnIndex] ??
            panic("A preflighted column deletion lost its coordinate");
          tr.setNodeAttribute(instruction.source.position + 1 + cellOffset, "cellMarker", {
            kind: "del",
            info: columnRevision,
          });
        }
        applied.push({ operand: edit.operand, revisionIds: Object.freeze([id]) });
      }
      const insertions = instruction.edits
        .filter(
          (edit): edit is Extract<ColumnEdit, { readonly type: "insert" }> =>
            edit.type === "insert",
        )
        .toSorted(
          (left, right) =>
            right.baseBoundaryIndex - left.baseBoundaryIndex ||
            right.targetColumnIndex - left.targetColumnIndex,
        );
      const ids = new Map<Extract<ColumnEdit, { readonly type: "insert" }>, number>();
      for (const edit of insertions) ids.set(edit, allocate());
      for (let rowIndex = map.height - 1; rowIndex >= 0; rowIndex--) {
        for (const edit of insertions) {
          const id = ids.get(edit) ?? panic("A column insertion lost its revision allocation");
          const cell =
            edit.targetCells[rowIndex] ??
            panic("A preflighted column insertion lost its target cell");
          const position =
            instruction.source.position +
            1 +
            map.positionAt(rowIndex, edit.baseBoundaryIndex, table);
          tr.insert(
            position,
            cell.type.create(
              {
                ...cell.attrs,
                cellMarker: { kind: "ins", info: structuralRevision(id, revision) },
              },
              cell.content,
              cell.marks,
            ),
          );
        }
      }
      for (const edit of insertions) {
        applied.push({
          operand: edit.operand,
          revisionIds: Object.freeze([
            ids.get(edit) ?? panic("A column insertion lost its revision allocation"),
          ]),
        });
      }
      if (instruction.trackedColumnWidths) {
        setTrackedTableGrid(tr, instruction.source.position, instruction.trackedColumnWidths);
      }
      markStructuralChange(tr);
      break;
    }
    default: {
      const unreachable: never = instruction;
      return panic("Unhandled preflighted table-structure instruction", {
        instruction: unreachable,
      });
    }
  }
  return {
    transaction: tr,
    nextRevisionId,
    applied: Object.freeze(applied),
    localPositionMappingSteps: 0,
  };
};
