/**
 * Section Break Commands
 * @packageDocumentation
 * @public
 */

import type { Node as PMNode } from "prosemirror-model";
import type { Command, Transaction } from "prosemirror-state";
import { TextSelection } from "prosemirror-state";

type InsertableSectionBreak = "nextPage" | "continuous";

/** What a paragraph states when a section ends at its mark. */
type SectionEndpointAttrs = {
  sectionBreakType: unknown;
  _sectionProperties: unknown;
};

const sectionEndpointOf = (paragraph: PMNode): SectionEndpointAttrs | null => {
  const sectionProperties = paragraph.attrs["_sectionProperties"];
  const sectionBreakType = paragraph.attrs["sectionBreakType"];
  if (sectionProperties == null && sectionBreakType == null) {
    return null;
  }
  return { sectionBreakType, _sectionProperties: sectionProperties };
};

/** The position of a paragraph's mark: the end of the content it closes. */
type ParagraphMarkEndpoint = {
  markPosition: number;
  endpoint: SectionEndpointAttrs;
};

const sectionEndingMarks = (doc: PMNode): ParagraphMarkEndpoint[] => {
  const marks: ParagraphMarkEndpoint[] = [];
  doc.descendants((node, position) => {
    if (node.type.name !== "paragraph") {
      return true;
    }
    const endpoint = sectionEndpointOf(node);
    if (endpoint) {
      marks.push({ markPosition: position + node.nodeSize - 1, endpoint });
    }
    return false;
  });
  return marks;
};

const restoreSectionEndingMarks = (
  transaction: Transaction,
  marks: readonly ParagraphMarkEndpoint[],
): void => {
  for (const { markPosition, endpoint } of marks) {
    const mapped = transaction.mapping.mapResult(markPosition, -1);
    if (mapped.deleted) {
      continue;
    }
    const $mark = transaction.doc.resolve(mapped.pos);
    const paragraph = $mark.parent;
    if ($mark.depth === 0 || paragraph.type.name !== "paragraph" || sectionEndpointOf(paragraph)) {
      continue;
    }
    transaction.setNodeMarkup($mark.before(), undefined, { ...paragraph.attrs, ...endpoint });
  }
};

/**
 * Run a command, keeping every section break on the paragraph whose mark
 * survived it.
 *
 * A section break is a property of a paragraph *mark*, and a join keeps the
 * trailing paragraph's mark: Backspace at the start of a section-ending
 * paragraph merges it into its predecessor, and the merged paragraph is the
 * one that still ends the section. ProseMirror's `join` keeps the *first*
 * node's attrs, so the break would go with the mark that was consumed rather
 * than stay with the mark that survived. Resolving a tracked paragraph-mark
 * deletion transfers it explicitly for the same reason.
 *
 * The transfer is keyed on the mark's position surviving the transaction, not
 * on which command ran, so deleting a section-ending paragraph outright still
 * removes its section: that mark is gone, and nothing inherits it.
 */
export const keepSectionBreaksOnSurvivingMarks =
  (command: Command): Command =>
  (state, dispatch, view) => {
    if (!dispatch) {
      return command(state, undefined, view);
    }
    const marks = sectionEndingMarks(state.doc);
    if (marks.length === 0) {
      return command(state, dispatch, view);
    }
    let captured: Transaction | null = null;
    const capture = (transaction: Transaction): void => {
      captured = transaction;
    };
    if (!command(state, capture, view)) {
      return false;
    }
    // Commands that report success without dispatching leave the document as
    // it was, so there is no mark to follow.
    const transaction: Transaction | null = captured;
    if (transaction === null) {
      return true;
    }
    restoreSectionEndingMarks(transaction, marks);
    dispatch(transaction);
    return true;
  };

/**
 * Insert a section break at the current cursor position.
 *
 * In OOXML a section break is the `sectPr` carried by the *last* paragraph of a
 * section. So we split the current paragraph at the cursor, mark the first half
 * with `sectionBreakType` (it becomes the section end), and leave the cursor in
 * the second half (the first paragraph of the new section). Content after the
 * cursor therefore flows into the new section — onto a new page for `nextPage`,
 * or in place for `continuous`.
 */
function insertSectionBreakAtCursor(breakType: InsertableSectionBreak): Command {
  return (state, dispatch) => {
    const { schema } = state;
    const paragraphType = schema.nodes["paragraph"];
    if (!paragraphType) return false;

    const { $from } = state.selection;
    // Section breaks only belong on top-level body paragraphs. Inside a table
    // cell or block SDT a `w:sectPr` is invalid OOXML, so refuse to act there
    // (the menu item becomes a no-op rather than corrupting the document).
    const isTopLevel = $from.parent.isTextblock
      ? state.doc.resolve($from.before()).depth === 0
      : $from.depth === 0;
    if (!isTopLevel) return false;

    if (dispatch) {
      const tr = state.tr;
      let cursorPos: number;

      if ($from.parent.isTextblock) {
        // Position of the paragraph node the cursor sits in. Unaffected by the
        // split below, since the split happens at a later position.
        const paraPos = $from.before();
        tr.split($from.pos);
        const firstPara = tr.doc.nodeAt(paraPos);
        if (firstPara) {
          tr.setNodeMarkup(paraPos, undefined, {
            ...firstPara.attrs,
            sectionBreakType: breakType,
          });
        }
        // The split position maps to the start of the second paragraph's
        // content (the first paragraph of the new section) — put the cursor there.
        cursorPos = tr.mapping.map($from.pos);
      } else {
        // Not in a textblock — insert a section-ending empty paragraph here.
        // Place the cursor *inside* the new paragraph (pos + 1). Mapping
        // `$from.pos` would land on a block boundary, which is not a valid
        // `TextSelection` position and would throw.
        const pos = $from.pos;
        tr.insert(pos, paragraphType.create({ sectionBreakType: breakType }));
        cursorPos = pos + 1;
      }

      tr.setSelection(TextSelection.create(tr.doc, cursorPos));
      dispatch(tr.scrollIntoView());
    }

    return true;
  };
}

/**
 * Insert a "next page" section break at the cursor — starts a new section on a
 * new page.
 */
export const insertSectionBreakNextPage: Command = insertSectionBreakAtCursor("nextPage");

/**
 * Insert a "continuous" section break at the cursor — starts a new section on
 * the same page.
 */
export const insertSectionBreakContinuous: Command = insertSectionBreakAtCursor("continuous");
