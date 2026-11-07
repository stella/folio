import type { Node as PMNode } from "prosemirror-model";
import type { Command, EditorState, Transaction } from "prosemirror-state";

import {
  type CanonicalCommandIntent,
  type CanonicalTableFormatTarget,
  withCanonicalCommand,
} from "./canonicalCommands";
import {
  tableAttrsToFormatting,
  tableCellAttrsToFormatting,
  tableRowAttrsToFormatting,
} from "./conversion/fromProseDoc";
import type { TableAttrs, TableCellAttrs, TableRowAttrs } from "./schema/nodes";

const SCOPES = {
  table: "table",
  tableRow: "row",
  tableCell: "cell",
  tableHeader: "cell",
} as const satisfies Record<string, CanonicalTableFormatTarget["scope"]>;
type ScopedNodeName = keyof typeof SCOPES;
const isScoped = (name: string): name is ScopedNodeName =>
  Object.hasOwn(SCOPES, name);

/** The formatting the save writes for a table, row or cell node. */
const formattingOf = (node: PMNode): Record<string, unknown> => {
  switch (SCOPES[node.type.name as ScopedNodeName]) {
    case "table":
      return { ...tableAttrsToFormatting(node.attrs as TableAttrs) };
    case "row":
      return { ...tableRowAttrsToFormatting(node.attrs as TableRowAttrs) };
    case "cell":
      return { ...tableCellAttrsToFormatting(node.attrs as TableCellAttrs) };
  }
};

/** Keys whose saved value changed: the new value, or null where the key went. */
const patchBetween = (
  before: Record<string, unknown>,
  after: Record<string, unknown>,
): Record<string, unknown> => {
  const patch: Record<string, unknown> = {};
  for (const key of new Set([...Object.keys(before), ...Object.keys(after)])) {
    if (JSON.stringify(before[key]) === JSON.stringify(after[key])) continue;
    patch[key] = Object.hasOwn(after, key) ? after[key] : null;
  }
  return patch;
};

/** A position inside the node's first paragraph, which addresses it. */
const firstParagraphPosition = (node: PMNode, position: number) => {
  let found: number | undefined;
  node.descendants((child, offset) => {
    if (found !== undefined) return false;
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
    const patch = patchBetween(formattingOf(node), formattingOf(after));
    if (Object.keys(patch).length === 0) return true;
    const at = firstParagraphPosition(node, position);
    if (at === undefined) {
      lowerable = false;
      return false;
    }
    targets.push({ scope: SCOPES[node.type.name], at, patch } as never);
    return true;
  });
  // Attribute changes alone must account for the whole command.
  if (!lowerable || !expected.doc.eq(changed.doc)) return undefined;
  return targets.length === 0 ? [] : [{ type: "formatTable", targets }];
};

/** Register a command whose whole effect is table, row and cell attributes. */
export const withCanonicalTableAttributeCommand = (command: Command): Command =>
  withCanonicalCommand(command, (state) =>
    lowerTableAttributeCommand(command, state),
  );
