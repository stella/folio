import { describe, expect, test } from "bun:test";

import { documentShape } from "../../__tests__/documentShapes";
import {
  createHarnessState,
  harnessManager,
  HeadlessEditorView,
  parseShapeDocument,
  placeSelection,
  readBack,
  saveHarnessState,
  summarizeState,
} from "../../__tests__/editorHarness";
import type { Document, Paragraph } from "../../types/document";

const paragraphHolding = (document: Document, text: string): Paragraph => {
  const found = document.package.document.content.find(
    (block): block is Paragraph =>
      block.type === "paragraph" && JSON.stringify(block).includes(text),
  );
  if (!found) {
    throw new Error(`No paragraph holds "${text}"`);
  }
  return found;
};

const indentAndSave = async (
  shapeId: string,
  focus: string,
  command: string,
  args: readonly unknown[] = [],
) => {
  const base = await parseShapeDocument(await documentShape(shapeId).build());
  const before = placeSelection(createHarnessState(base, "editing"), focus, "caret-middle");
  if (!before) {
    throw new Error("the shape has no focus paragraph");
  }
  const view = new HeadlessEditorView(before);
  expect(harnessManager().requireCommand(command)(...args)(view.state, view.dispatch)).toBe(true);
  const saved = await saveHarnessState(view.state, base);
  return { view, saved, reopened: await parseShapeDocument(saved.bytes) };
};

describe("paragraph indentation on save", () => {
  test.each([
    // A paragraph whose source states `w:bidi` and `w:jc` but no `w:ind`.
    ["rtl-cjk", "مرحبا", 720],
    // A paragraph whose source states only `w:jc` (under a tracked change).
    ["pending-property-change", "Centred by a tracked change.", 720],
    // A list item whose source states `w:ind` itself.
    ["single-decimal-list", "Beta", 1440],
  ] as const)(
    "an indent the editor applied to a paragraph with its own w:pPr is saved (%s)",
    async (shapeId, focus, expected) => {
      const { reopened } = await indentAndSave(shapeId, focus, "increaseIndent");

      expect(paragraphHolding(reopened, focus).formatting?.indentLeft).toBe(expected);
    },
  );

  test("the block snapshot reports an indentation a command set, as the saved package does", async () => {
    const { view, saved } = await indentAndSave(
      "plain-markdown",
      "First item",
      "setIndentLeft",
      [720],
    );

    const live = summarizeState(view.state).find((block) => block["text"] === "First item");
    const reopened = (await readBack(saved.bytes)).summary.find(
      (block) => block["text"] === "First item",
    );
    expect(live?.["directIndentation"]).toEqual({ indentLeft: 720 });
    expect(reopened?.["directIndentation"]).toEqual(live?.["directIndentation"]);
  });

  test("a hanging indent reads with a negative first line, as the model states it", async () => {
    const { view, reopened } = await indentAndSave("rtl-cjk", "مرحبا", "setIndentFirstLine", [
      360,
      true,
    ]);

    const live = summarizeState(view.state).find((block) =>
      String(block["text"]).startsWith("مرحبا"),
    );
    expect(live?.["directIndentation"]).toEqual({ indentFirstLine: -360, hangingIndent: true });
    expect(paragraphHolding(reopened, "مرحبا").formatting).toMatchObject({
      indentFirstLine: -360,
      hangingIndent: true,
    });
  });
});
