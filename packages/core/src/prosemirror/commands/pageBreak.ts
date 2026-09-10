/**
 * Page Break Commands
 */

import type { Command } from "prosemirror-state";

import { pageBreakRunParagraphProjectionDisposition } from "../pageBreakRunProjection";

const UNSUPPORTED_PAGE_BREAK_RUN_ANCESTORS = new Set(["tableCell", "tableHeader", "textBox"]);

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
    pageBreakRunParagraphProjectionDisposition($from.parent).status === "unsupported"
  ) {
    return false;
  }

  if (dispatch) {
    const pageBreak = pageBreakRunType.create();
    dispatch(state.tr.replaceSelectionWith(pageBreak).scrollIntoView());
  }

  return true;
};
