/**
 * Plugins read the document's numbering while they apply a transaction, when
 * only the plugins before them hold their new state. The editor states the
 * editor builds place the numbering plugin first, so what they read is the
 * package's numbering, not one rebuilt from the paragraphs alone.
 */
import { expect, spyOn, test } from "bun:test";
import { type EditorState, TextSelection } from "prosemirror-state";

import { paragraphNumberingReference } from "@stll/docx-core/model";

import { fromMarkdown } from "../markdown/fromMarkdown";
import * as listInstanceReferences from "../prosemirror/listInstanceReferences";
import {
  extractSelectionContext,
  getSelectionContext,
} from "../prosemirror/plugins/selectionTracker";
import { singletonManager } from "../prosemirror/schema";
import { createHiddenEditorState } from "./hiddenEditorManager";

/** Two paragraphs; the second is a bullet item the package alone defines. */
const bulletedDocument = () => {
  const document = fromMarkdown("Opening words\n\nBulleted words");
  document.package.numbering = {
    abstractNums: [
      {
        abstractNumId: 3,
        multiLevelType: "singleLevel",
        levels: [
          {
            ilvl: 0,
            start: 1,
            numFmt: "bullet",
            lvlText: "•",
            suffix: "tab",
          },
        ],
      },
    ],
    nums: [{ numId: 3, abstractNumId: 3 }],
  };
  const bulleted = document.package.document.content[1];
  if (bulleted?.type !== "paragraph")
    throw new TypeError("Expected the second block to be a paragraph.");
  bulleted.formatting = {
    ...bulleted.formatting,
    numPr: paragraphNumberingReference({ numId: 3, ilvl: 0 }),
  };
  return document;
};

const caretInSecondParagraph = (state: EditorState): EditorState => {
  let at = -1;
  let seen = 0;
  state.doc.descendants((node, position) => {
    if (node.type.name !== "paragraph") return true;
    if (seen++ === 1) at = position + 2;
    return false;
  });
  return state.apply(
    state.tr.setSelection(TextSelection.create(state.doc, at)),
  );
};

test("typing in a list item reads the package's numbering without rebuilding it", () => {
  let state = caretInSecondParagraph(
    createHiddenEditorState({
      document: bulletedDocument(),
      manager: singletonManager,
    }),
  );
  // Typing changes no numbering input, so nothing completes the numbering
  // again: no plugin rebuilds it from the paragraphs.
  const completions = spyOn(listInstanceReferences, "completeNumberingForDoc");
  try {
    state = state.apply(state.tr.insertText("x"));
    expect(completions).not.toHaveBeenCalled();
  } finally {
    completions.mockRestore();
  }
  const expected = { type: "bullet", level: 0, numId: 3 };
  expect(extractSelectionContext(state).listState).toEqual(expected);
  expect(getSelectionContext(state)?.listState).toEqual(expected);
});
