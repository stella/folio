import { describe, expect, test } from "bun:test";
import type { EditorState } from "prosemirror-state";

import { documentShape } from "../__tests__/documentShapes";
import { CONFORMANCE_OPERATIONS } from "../__tests__/editorCommandConformance";
import {
  createHarnessState,
  HeadlessEditorView,
  parseShapeDocument,
  placeSelection,
  readBack,
  saveHarnessState,
  summarizeState,
} from "../__tests__/editorHarness";

/**
 * A hyperlink's color and underline come from its character style, not from
 * direct formatting. Clearing the direct color or applying a paragraph style
 * resets direct formatting only, so the link keeps painting as a link — and
 * the editor agrees with the document it saves.
 */
const linkRunMarks = (state: EditorState): string[] => {
  let marks: string[] = [];
  state.doc.descendants((node) => {
    if (node.isText && node.text === "site") {
      marks = node.marks.map((mark) => mark.type.name).toSorted();
    }
  });
  return marks;
};

/** Run a matrix operation (as the host drives it) over the paragraph holding the link. */
const runOverLink = async (operationId: string) => {
  const operation = CONFORMANCE_OPERATIONS.find(({ id }) => id === operationId);
  if (!operation) {
    throw new Error(`No operation ${operationId}`);
  }
  const shape = documentShape("fields-links-bookmarks");
  const base = await parseShapeDocument(await shape.build());
  const state = placeSelection(createHarnessState(base, "editing"), shape.focus, "paragraph");
  if (!state) {
    throw new Error("No selection");
  }
  const before = linkRunMarks(state);
  const view = new HeadlessEditorView(state);
  expect(operation.run({ view, base, focus: shape.focus })).not.toBe(false);
  const { bytes } = await saveHarnessState(view.state, base);
  return { before, after: view.state, reopened: (await readBack(bytes)).summary };
};

describe("style formatting survives a reset of direct formatting", () => {
  test("clearing the text color keeps a hyperlink's style color", async () => {
    const { before, after, reopened } = await runOverLink("command:clearTextColor");
    expect(linkRunMarks(after)).toEqual(before);
    expect(summarizeState(after)).toEqual(reopened);
  });

  test("applying a paragraph style keeps a hyperlink's character style formatting", async () => {
    const { after, reopened } = await runOverLink("command:applyStyle(Heading1)");
    expect(linkRunMarks(after)).toContain("textColor");
    expect(linkRunMarks(after)).toContain("underline");
    expect(summarizeState(after)).toEqual(reopened);
  });
});
