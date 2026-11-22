/**
 * Section breaks a paste or a drop would put inside a table cell.
 *
 * A section ends at a paragraph outside every table, so a paragraph that lands
 * in a cell, or that stands in a pasted cell, comes without the break it ended
 * its own section with.
 */

import { Fragment, type Node as PMNode, Slice } from "prosemirror-model";
import type { Selection, Transaction } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";

import type { SectionProperties } from "../../../types/document";
import { sectionPropertiesOf } from "../../sectionCarrier";

const isCell = (node: PMNode): boolean => {
  const role = node.type.spec["tableRole"];
  return role === "cell" || role === "header_cell";
};

const isInCell = ({ $from }: Selection): boolean => {
  if ($from.nodeAfter !== null && isCell($from.nodeAfter)) return true;
  for (let depth = $from.depth; depth > 0; depth -= 1)
    if (isCell($from.node(depth))) return true;
  return false;
};

const withoutSections = (node: PMNode, inCell: boolean): PMNode => {
  if (node.type.name === "paragraph")
    return inCell && sectionPropertiesOf(node) !== null
      ? node.type.create(
          { ...node.attrs, _sectionProperties: null },
          node.content,
          node.marks,
        )
      : node;
  if (node.isLeaf || node.isTextblock) return node;
  const content = withoutSectionsIn(node.content, inCell || isCell(node));
  return content === node.content ? node : node.copy(content);
};

const withoutSectionsIn = (fragment: Fragment, inCell: boolean): Fragment => {
  let changed = false;
  const nodes: PMNode[] = [];
  // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Fragment.forEach
  fragment.forEach((node) => {
    const next = withoutSections(node, inCell);
    changed ||= next !== node;
    nodes.push(next);
  });
  return changed ? Fragment.from(nodes) : fragment;
};

type CellSection = { position: number; record: SectionProperties };

/** The paragraphs in table cells that end a section, in document order. */
const cellSectionsIn = (doc: PMNode): CellSection[] => {
  const found: CellSection[] = [];
  const visit = (parent: PMNode, start: number, inCell: boolean): void => {
    // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
    parent.forEach((node, offset) => {
      const position = start + offset;
      if (node.type.name === "paragraph") {
        const record = inCell ? sectionPropertiesOf(node) : null;
        if (record !== null) found.push({ position, record });
        return;
      }
      if (!node.isLeaf && !node.isTextblock)
        visit(node, position + 1, inCell || isCell(node));
    });
  };
  visit(doc, 0, false);
  return found;
};

/**
 * Where a drop put a section break in a table cell. The drop's target is not
 * the selection its slice was read in, so the slice still carries the breaks
 * of the paragraphs dropped.
 */
export const droppedCellSections = (
  transactions: readonly Transaction[],
  before: PMNode,
  after: PMNode,
): number[] => {
  if (
    !transactions.some(
      (transaction) =>
        transaction.docChanged && transaction.getMeta("uiEvent") === "drop",
    )
  )
    return [];
  const held = new Set(cellSectionsIn(before).map(({ record }) => record));
  return cellSectionsIn(after)
    .filter(({ record }) => !held.has(record))
    .map(({ position }) => position);
};

/**
 * A pasted slice without the section breaks its paragraphs would end in a
 * cell: in its own cells, and everywhere when it replaces a selection in one.
 * A drag lands elsewhere; {@link droppedCellSections} reads where.
 */
export const withoutPastedCellSections = (
  slice: Slice,
  view: Pick<EditorView, "dragging" | "state">,
): Slice => {
  const content = withoutSectionsIn(
    slice.content,
    view.dragging === null && isInCell(view.state.selection),
  );
  return content === slice.content
    ? slice
    : new Slice(content, slice.openStart, slice.openEnd);
};
