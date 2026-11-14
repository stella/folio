import type { Node as PMNode } from "prosemirror-model";
import type { Command, EditorState, Transaction } from "prosemirror-state";
import { isInTable, selectedRect, type TableRect } from "prosemirror-tables";

import type { Table, TableRow } from "@stll/docx-core/model";
import { DOCUMENT_OP_TYPES, type FormattingPatch } from "@stll/docx-core/ops";

import {
  expectTableAttrs,
  expectTableCellAttrs,
  expectTableRowAttrs,
} from "./attrs";
import {
  type CanonicalCommandIntent,
  type CanonicalTableFormatTarget,
  withCanonicalCommand,
} from "./canonicalCommands";
import {
  standaloneTableRowsFromProseMirror,
  tableAttrsToFormatting,
  tableCellAttrsToFormatting,
  tableRowAttrsToFormatting,
} from "./conversion/fromProseDoc";

const SCOPES = {
  table: "table",
  tableRow: "row",
  tableCell: "cell",
  tableHeader: "cell",
} as const satisfies Record<string, CanonicalTableFormatTarget["scope"]>;
type ScopedNodeName = keyof typeof SCOPES;
const isScoped = (name: string): name is ScopedNodeName =>
  Object.hasOwn(SCOPES, name);

/** Keys whose saved value changed: the new value, or null where the key went. */
const patchBetween = <Formatting extends object>(
  before: Formatting | undefined,
  after: Formatting | undefined,
): FormattingPatch<Formatting> => {
  const patch: {
    -readonly [Key in keyof Formatting]?: Formatting[Key] | null;
  } = {};
  // Object.keys widens to string; these are the formatting's own keys.
  const keys = Object.keys({ ...before, ...after }) as (keyof Formatting)[];
  for (const key of keys) {
    const was = before?.[key];
    const is = after?.[key];
    if (JSON.stringify(was) === JSON.stringify(is)) continue;
    patch[key] = after !== undefined && Object.hasOwn(after, key) ? is : null;
  }
  return patch;
};

/**
 * The patch of the formatting the save writes for a table, row or cell node,
 * between its attributes before and after the command, as a target at `at`.
 */
const targetBetween = (
  scope: CanonicalTableFormatTarget["scope"],
  before: PMNode,
  after: PMNode,
  at: number,
): CanonicalTableFormatTarget => {
  switch (scope) {
    case "table":
      return {
        scope,
        at,
        patch: patchBetween(
          tableAttrsToFormatting(expectTableAttrs(before)),
          tableAttrsToFormatting(expectTableAttrs(after)),
        ),
      };
    case "row":
      return {
        scope,
        at,
        patch: patchBetween(
          tableRowAttrsToFormatting(expectTableRowAttrs(before)),
          tableRowAttrsToFormatting(expectTableRowAttrs(after)),
        ),
      };
    case "cell":
      return {
        scope,
        at,
        patch: patchBetween(
          tableCellAttrsToFormatting(expectTableCellAttrs(before)),
          tableCellAttrsToFormatting(expectTableCellAttrs(after)),
        ),
      };
  }
};

/**
 * A position inside the first paragraph that belongs to the node itself, which
 * addresses it. A paragraph of a nested table would address the inner table.
 */
const firstParagraphPosition = (node: PMNode, position: number) => {
  let found: number | undefined;
  node.descendants((child, offset) => {
    if (found !== undefined || child.type.name === "table") return false;
    if (child.isTextblock) {
      found = position + 1 + offset + 1;
      return false;
    }
    return true;
  });
  return found;
};

/**
 * What an attribute-only table command does, as table formatting patches.
 *
 * The command runs against the state without dispatching to the view. Every
 * table, row and cell whose attributes it changed lowers to a patch of the
 * formatting the save would write for it, so the session applies exactly what
 * the editor would have saved. A command that changes anything else (content,
 * a node's type, a span) has no attribute lowering and stays refused.
 */
export const lowerTableAttributeCommand = (
  command: Command,
  state: EditorState,
): readonly CanonicalCommandIntent[] | undefined => {
  let transaction: Transaction | undefined;
  if (
    !command(state, (dispatched) => {
      transaction = dispatched;
    }) ||
    transaction === undefined
  )
    return undefined;
  const changed = transaction;
  const targets: CanonicalTableFormatTarget[] = [];
  const grids: CanonicalCommandIntent[] = [];
  const expected = state.tr;
  let lowerable = true;
  state.doc.descendants((node, position) => {
    if (!lowerable) return false;
    if (!isScoped(node.type.name)) return true;
    // A markup change replaces the node's opening token; its start maps backward.
    const after = changed.doc.nodeAt(changed.mapping.map(position, -1));
    if (after === null || after.type !== node.type) {
      lowerable = false;
      return false;
    }
    if (
      after.attrs["colspan"] !== node.attrs["colspan"] ||
      after.attrs["rowspan"] !== node.attrs["rowspan"]
    ) {
      lowerable = false;
      return false;
    }
    if (JSON.stringify(after.attrs) === JSON.stringify(node.attrs)) return true;
    expected.setNodeMarkup(position, undefined, after.attrs, node.marks);
    const at = firstParagraphPosition(node, position);
    if (node.type.name === "table") {
      const grid = gridEdit(node, after, at);
      if (grid === null) {
        lowerable = false;
        return false;
      }
      if (grid !== undefined) grids.push(grid);
    }
    const target = targetBetween(SCOPES[node.type.name], node, after, at ?? 0);
    if (Object.keys(target.patch).length === 0) return true;
    if (at === undefined) {
      lowerable = false;
      return false;
    }
    targets.push(target);
    return true;
  });
  // Attribute changes alone must account for the whole command.
  if (!lowerable || !expected.doc.eq(changed.doc)) return undefined;
  return [
    ...grids,
    ...(targets.length === 0
      ? []
      : [{ type: "formatTable", targets } as const]),
  ];
};

/**
 * The grid a command left on a table, as a grid edit: `undefined` when the
 * grid is unchanged, `null` when the change cannot be written as a grid.
 */
const gridEdit = (
  before: PMNode,
  after: PMNode,
  at: number | undefined,
): CanonicalCommandIntent | null | undefined => {
  const was = expectTableAttrs(before).columnWidths;
  const is = expectTableAttrs(after).columnWidths;
  if (JSON.stringify(was) === JSON.stringify(is)) return undefined;
  if (at === undefined || is === undefined || is.length === 0) return null;
  return {
    type: "editTable",
    at,
    edit: { type: DOCUMENT_OP_TYPES.SET_TABLE_GRID, columnWidths: [...is] },
  };
};

/** Register a command whose whole effect is table, row and cell attributes. */
export const withCanonicalTableAttributeCommand = (command: Command): Command =>
  withCanonicalCommand(command, (state) =>
    lowerTableAttributeCommand(command, state),
  );

/**
 * Register a command whose whole effect is the selection, such as selecting a
 * table, a row or a column: it runs as it is and changes no content. A
 * command that would change the document has no such description.
 */
export const withCanonicalSelectionCommand = (command: Command): Command =>
  withCanonicalCommand(command, (state) => {
    let transaction: Transaction | undefined;
    if (
      !command(state, (dispatched) => {
        transaction = dispatched;
      }) ||
      transaction === undefined ||
      transaction.docChanged
    )
      return undefined;
    return [{ type: "selection" }];
  });

/** A paragraph of the table's own, outside the given columns, that addresses it. */
export const tableAnchor = (
  rect: TableRect,
  avoid: { left: number; right: number } = { left: 0, right: 0 },
): number | undefined => {
  const { map, table, tableStart } = rect;
  for (let row = 0; row < map.height; row += 1)
    for (let column = 0; column < map.width; column += 1) {
      if (column >= avoid.left && column < avoid.right) continue;
      const offset = map.map[row * map.width + column];
      if (offset === undefined) continue;
      const cell = table.nodeAt(offset);
      // A row's gridBefore and gridAfter fillers have no source paragraph.
      if (cell === null || cell.attrs["_omittedGridSlot"] != null) continue;
      const cellRect = map.findCell(offset);
      if (cellRect.left < avoid.right && cellRect.right > avoid.left) continue;
      const at = firstParagraphPosition(cell, tableStart + offset);
      if (at !== undefined) return at;
    }
  return undefined;
};

type StructureLowering = (
  rect: TableRect,
  state: EditorState,
  changed: Transaction,
) => readonly CanonicalCommandIntent[] | undefined;

/**
 * Register a command that changes a table's columns or cells, lowered from the
 * selection it acts on to the engine's table edit. The command still runs
 * against the state first: one it would refuse has no lowering.
 */
const withCanonicalTableStructureCommand = (
  command: Command,
  lower: StructureLowering,
): Command =>
  withCanonicalCommand(command, (state) => {
    if (!isInTable(state)) return undefined;
    let transaction: Transaction | undefined;
    if (
      !command(state, (dispatched) => {
        transaction = dispatched;
      }) ||
      transaction === undefined
    )
      return undefined;
    return lower(selectedRect(state), state, transaction);
  });

/** The width the command gave the column it inserted, from the table it left. */
const insertedWidth = (
  rect: TableRect,
  changed: Transaction,
  column: number,
): number | undefined => {
  const after = changed.doc.nodeAt(
    changed.mapping.map(rect.tableStart - 1, -1),
  );
  if (after === null || after.type.name !== "table") return undefined;
  return expectTableAttrs(after).columnWidths?.[column];
};

const insertColumn =
  (side: "left" | "right"): StructureLowering =>
  (rect, _state, changed) => {
    const column = side === "left" ? rect.left : rect.right;
    const width = insertedWidth(rect, changed, column);
    const at = tableAnchor(rect);
    if (width === undefined || at === undefined) return undefined;
    return [
      {
        type: "editTable",
        at,
        edit: { type: DOCUMENT_OP_TYPES.INSERT_COLUMN, column, width },
      },
    ];
  };

/** Each selected column, from the right so the indices still hold. */
const deleteColumns: StructureLowering = (rect) => {
  const at = tableAnchor(rect, { left: rect.left, right: rect.right });
  if (at === undefined) return undefined;
  const intents: CanonicalCommandIntent[] = [];
  for (let column = rect.right - 1; column >= rect.left; column -= 1)
    intents.push({
      type: "editTable",
      at,
      edit: { type: DOCUMENT_OP_TYPES.DELETE_COLUMN, column },
    });
  return intents;
};

const mergeCells: StructureLowering = (rect) => {
  const at = tableAnchor(rect);
  if (at === undefined) return undefined;
  return [
    {
      type: "editTable",
      at,
      edit: {
        type: DOCUMENT_OP_TYPES.MERGE_CELLS,
        top: rect.top,
        bottom: rect.bottom,
        left: rect.left,
        right: rect.right,
      },
    },
  ];
};

/** The cell at the selection's top left, addressed by its own paragraph. */
const splitCell: StructureLowering = (rect) => {
  const offset = rect.map.map[rect.top * rect.map.width + rect.left];
  const cell = offset === undefined ? null : rect.table.nodeAt(offset);
  if (offset === undefined || cell === null) return undefined;
  const at = firstParagraphPosition(cell, rect.tableStart + offset);
  if (at === undefined) return undefined;
  return [
    { type: "editTable", at, edit: { type: DOCUMENT_OP_TYPES.SPLIT_CELL } },
  ];
};

export const withCanonicalColumnInsertion = (
  command: Command,
  side: "left" | "right",
): Command => withCanonicalTableStructureCommand(command, insertColumn(side));
export const withCanonicalColumnDeletion = (command: Command): Command =>
  withCanonicalTableStructureCommand(command, deleteColumns);
export const withCanonicalCellMerge = (command: Command): Command =>
  withCanonicalTableStructureCommand(command, mergeCells);
export const withCanonicalCellSplit = (command: Command): Command =>
  withCanonicalTableStructureCommand(command, splitCell);

/** A paragraph of the table's own in the given row. */
export const rowAnchor = (rect: TableRect, row: number): number | undefined => {
  const { map, table, tableStart } = rect;
  for (let column = 0; column < map.width; column += 1) {
    const offset = map.map[row * map.width + column];
    if (offset === undefined) continue;
    const cell = table.nodeAt(offset);
    if (cell === null || cell.attrs["_omittedGridSlot"] != null) continue;
    // A cell reaching down from an earlier row belongs to that row.
    if (map.findCell(offset).top !== row) continue;
    const at = firstParagraphPosition(cell, tableStart + offset);
    if (at !== undefined) return at;
  }
  return undefined;
};

/**
 * The row the command inserted, as the save would write it. Revision marks
 * the command put on it are left to the engine, which records its own.
 */
const insertedRow = (
  rect: TableRect,
  changed: Transaction,
  index: number,
): TableRow | undefined => {
  const after = changed.doc.nodeAt(
    changed.mapping.map(rect.tableStart - 1, -1),
  );
  if (after === null || after.type.name !== "table") return undefined;
  const row = standaloneTableRowsFromProseMirror(after)[index];
  if (row === undefined) return undefined;
  const {
    structuralChange: _structuralChange,
    propertyChanges: _propertyChanges,
    ...plain
  } = row;
  return plain;
};

const insertRow =
  (side: "above" | "below"): StructureLowering =>
  (rect, _state, changed) => {
    const index = side === "above" ? rect.top : rect.bottom;
    const row = insertedRow(rect, changed, index);
    const at = rowAnchor(rect, side === "above" ? rect.top : rect.bottom - 1);
    if (row === undefined || at === undefined) return undefined;
    return [
      {
        type: "editTableRow",
        at,
        edit: { type: DOCUMENT_OP_TYPES.INSERT_ROW, at: index, row },
      },
    ];
  };

/** Each selected row from the bottom, addressed by a paragraph of its own. */
const deleteRows: StructureLowering = (rect) => {
  const intents: CanonicalCommandIntent[] = [];
  for (let row = rect.bottom - 1; row >= rect.top; row -= 1) {
    const at = rowAnchor(rect, row);
    if (at === undefined) return undefined;
    intents.push({
      type: "editTableRow",
      at,
      edit: { type: DOCUMENT_OP_TYPES.DELETE_ROW },
    });
  }
  return intents;
};

export const withCanonicalRowInsertion = (
  command: Command,
  side: "above" | "below",
): Command => withCanonicalTableStructureCommand(command, insertRow(side));
export const withCanonicalRowDeletion = (command: Command): Command =>
  withCanonicalTableStructureCommand(command, deleteRows);

/** The table the selection is in, addressed by a paragraph of its own. */
const deleteTable: StructureLowering = (rect) => {
  const at = tableAnchor(rect);
  if (at === undefined) return undefined;
  return [
    {
      type: "editTableBlock",
      at,
      edit: { type: DOCUMENT_OP_TYPES.DELETE_TABLE },
    },
  ];
};

export const withCanonicalTableDeletion = (command: Command): Command =>
  withCanonicalTableStructureCommand(command, deleteTable);

/**
 * Register the insert table command, lowered to a table and an empty
 * paragraph after the paragraph that holds the caret, as the command places
 * them. The table is the one the command built, without the revision marks
 * it put on its rows; in suggesting mode the engine records its own.
 */
export const withCanonicalTableInsertion = (command: Command): Command =>
  withCanonicalCommand(command, (state) => {
    const { $from } = state.selection;
    // Where the command puts the table: after the nearest paragraph or table
    // holding the selection's start, even from inside an inline content
    // control; at the story's start when nothing holds it (a selection of
    // the whole story). The table anchors to the paragraph it follows, or
    // goes before the paragraph that follows the table or starts the story.
    let depth = $from.depth;
    while (
      depth > 0 &&
      $from.node(depth).type.name !== "paragraph" &&
      $from.node(depth).type.name !== "table"
    )
      depth -= 1;
    const afterTable = depth > 0 && $from.node(depth).type.name === "table";
    const atStart =
      depth === 0 &&
      $from.pos === 0 &&
      state.doc.firstChild?.type.name === "paragraph";
    if (depth === 0 && !atStart) return undefined;
    const following = afterTable
      ? state.doc.resolve($from.after(depth)).nodeAfter
      : null;
    if (afterTable && following?.type.name !== "paragraph") return undefined;
    let transaction: Transaction | undefined;
    if (
      !command(state, (dispatched) => {
        transaction = dispatched;
      }) ||
      transaction === undefined
    )
      return undefined;
    // After a table the command writes an empty paragraph ahead of the new
    // table, so the two never touch.
    const at = atStart ? 0 : transaction.mapping.map($from.after(depth), -1);
    const leading = afterTable ? transaction.doc.nodeAt(at) : null;
    const inserted = transaction.doc.nodeAt(
      leading?.type.name === "paragraph" ? at + leading.nodeSize : at,
    );
    const before = atStart || afterTable;
    // Inside the paragraph the table goes after, or before: the story's
    // first, or the one after the table.
    const anchorPosition = (): number => {
      if (atStart) return 1;
      if (afterTable) return $from.after(depth) + 1;
      return $from.start(depth);
    };
    if (inserted === null || inserted.type.name !== "table") return undefined;
    const formatting = tableAttrsToFormatting(expectTableAttrs(inserted));
    const columnWidths = expectTableAttrs(inserted).columnWidths;
    const table: Table = {
      type: "table",
      ...(formatting === undefined ? {} : { formatting }),
      ...(columnWidths === undefined
        ? {}
        : { columnWidths: [...columnWidths] }),
      rows: standaloneTableRowsFromProseMirror(inserted).map(
        ({ structuralChange: _change, propertyChanges: _changes, ...row }) =>
          row,
      ),
    };
    return [
      {
        type: "editTableBlock",
        at: anchorPosition(),
        edit: {
          type: DOCUMENT_OP_TYPES.INSERT_TABLE,
          table,
          ...(before ? { side: "before" as const } : {}),
          ...(afterTable ? { separated: true } : {}),
        },
      },
    ];
  });
