import type { Node as PMNode } from "prosemirror-model";
import type { Command, EditorState, Transaction } from "prosemirror-state";

import type { FormattingPatch } from "@stll/docx-core/ops";

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
  return targets.length === 0 ? [] : [{ type: "formatTable", targets }];
};

/** Register a command whose whole effect is table, row and cell attributes. */
export const withCanonicalTableAttributeCommand = (command: Command): Command =>
  withCanonicalCommand(command, (state) =>
    lowerTableAttributeCommand(command, state),
  );
