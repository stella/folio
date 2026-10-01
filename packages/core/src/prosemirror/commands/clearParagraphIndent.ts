import { TextSelection, type Command } from "prosemirror-state";

import { expectParagraphAttrs } from "../attrs";

/** Clear indentation at paragraph start before Backspace joins paragraphs. */
export const clearIndentOnBackspace: Command = (state, dispatch) => {
  if (!(state.selection instanceof TextSelection)) {
    return false;
  }
  const { $cursor } = state.selection;
  if (!$cursor || $cursor.parentOffset !== 0 || $cursor.parent.type.name !== "paragraph") {
    return false;
  }
  const attrs = expectParagraphAttrs($cursor.parent);
  const hasFirstLine = typeof attrs.indentFirstLine === "number" && attrs.indentFirstLine > 0;
  const hasIndentLeft = typeof attrs.indentLeft === "number" && attrs.indentLeft > 0;
  if (!hasFirstLine && !attrs.hangingIndent && !hasIndentLeft) {
    return false;
  }
  if (dispatch) {
    const position = $cursor.before();
    const tr = state.tr.setNodeMarkup(position, undefined, {
      ...attrs,
      indentFirstLine: null,
      hangingIndent: null,
      indentLeft: null,
    });
    dispatch(tr.scrollIntoView());
  }
  return true;
};
