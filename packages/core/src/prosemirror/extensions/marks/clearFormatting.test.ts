import { describe, expect, test } from "bun:test";
import type { EditorState } from "prosemirror-state";

import { documentShape } from "../../../__tests__/documentShapes";
import {
  createHarnessState,
  findTextblock,
  HeadlessEditorView,
  parseShapeDocument,
  placeSelection,
} from "../../../__tests__/editorHarness";
import { clearFormatting } from "./markUtils";

/** The marks on the inline content of the paragraph holding `focus`. */
const paragraphMarkNames = (state: EditorState, focus: string): Set<string> => {
  const paragraph = findTextblock(state.doc, focus);
  if (!paragraph) {
    throw new Error(`No paragraph holds "${focus}"`);
  }
  const names = new Set<string>();
  paragraph.node.descendants((node) => {
    for (const mark of node.marks) {
      names.add(mark.type.name);
    }
    return true;
  });
  return names;
};

const clearParagraph = async (shapeId: string, focus: string) => {
  const document = await parseShapeDocument(await documentShape(shapeId).build());
  const before = placeSelection(createHarnessState(document, "editing"), focus, "paragraph");
  if (!before) {
    throw new Error("the shape has no focus paragraph");
  }
  const view = new HeadlessEditorView(before);
  expect(clearFormatting(view.state, view.dispatch)).toBe(true);
  return {
    before: paragraphMarkNames(before, focus),
    after: paragraphMarkNames(view.state, focus),
  };
};

describe("clearFormatting", () => {
  test.each([
    ["comments", "Commented words sit in this paragraph.", ["comment"]],
    ["fields-links-bookmarks", "See the site and the target.", ["hyperlink"]],
    ["notes", "A sentence with a footnote", ["footnoteRef"]],
    ["tracked-changes", "Kept text inserted text", ["insertion", "deletion"]],
  ] as const)("keeps what is not character formatting (%s)", async (shapeId, focus, kept) => {
    const { before, after } = await clearParagraph(shapeId, focus);

    for (const name of kept) {
      expect(before.has(name)).toBe(true);
      expect(after.has(name)).toBe(true);
    }
  });

  test("removes character formatting", async () => {
    const { before, after } = await clearParagraph("tracked-changes", "Bold by a tracked change.");

    expect(before.has("bold")).toBe(true);
    expect(after.has("bold")).toBe(false);
  });
});
