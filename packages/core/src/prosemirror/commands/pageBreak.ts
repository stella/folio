/**
 * Page Break Commands
 */

import type { Node as PMNode } from "prosemirror-model";
import type { Command } from "prosemirror-state";

import { pageBreakRunParagraphProjectionDisposition } from "../pageBreakRunProjection";

const UNSUPPORTED_PAGE_BREAK_RUN_ANCESTORS = new Set(["tableCell", "tableHeader", "textBox"]);

const hasTextBoxAnchorAtOrAfter = (
  paragraph: PMNode,
  contentStart: number,
  position: number,
): boolean => {
  let found = false;
  paragraph.descendants((descendant, offset) => {
    if (descendant.type.name === "textBoxAnchor" && contentStart + offset >= position) {
      found = true;
      return false;
    }
    return !found;
  });
  return found;
};

/**
 * Insert a page-break run at the current text cursor.
 *
 * Keeping the atom inline preserves its authored run, revision marks, and exact
 * position among neighbouring run content. Block-level `pageBreak` remains a
 * legacy input boundary and is never emitted by this command.
 */
export const insertPageBreak: Command = (state, dispatch) => {
  const { schema } = state;
  const pageBreakRunType = schema.nodes["pageBreakRun"];
  const { $from } = state.selection;
  if (!pageBreakRunType || !state.selection.empty || !$from.parent.isTextblock) {
    return false;
  }
  for (let depth = $from.depth; depth >= 0; depth--) {
    if (UNSUPPORTED_PAGE_BREAK_RUN_ANCESTORS.has($from.node(depth).type.name)) {
      return false;
    }
  }
  if (
    $from.parent.type.name !== "paragraph" ||
    pageBreakRunParagraphProjectionDisposition($from.parent).status === "approximate" ||
    // The break is about to land here, so ask the question of the paragraph
    // this would produce: an anchor the caret precedes would end up after it.
    hasTextBoxAnchorAtOrAfter($from.parent, $from.start(), $from.pos)
  ) {
    return false;
  }

  if (dispatch) {
    const pageBreak = pageBreakRunType.create();
    dispatch(state.tr.replaceSelectionWith(pageBreak).scrollIntoView());
  }

  return true;
};
