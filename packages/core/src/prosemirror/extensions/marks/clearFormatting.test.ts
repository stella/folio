import { describe, expect, test } from "bun:test";
import type { EditorState } from "prosemirror-state";

import { documentShape } from "../../../__tests__/documentShapes";
import {
  createHarnessState,
  HeadlessEditorView,
  parseShapeDocument,
  placeSelection,
} from "../../../__tests__/editorHarness";
import { clearFormatting } from "./markUtils";

const markNames = (state: EditorState): Set<string> => {
  const names = new Set<string>();
  state.doc.descendants((node) => {
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
  return { before, after: view.state };
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
      expect(markNames(before).has(name)).toBe(true);
      expect(markNames(after).has(name)).toBe(true);
    }
  });

  test("removes character formatting", async () => {
    const { before, after } = await clearParagraph("tracked-changes", "Bold by a tracked change.");

    expect(markNames(before).has("bold")).toBe(true);
    expect(markNames(after).has("bold")).toBe(false);
  });
});
