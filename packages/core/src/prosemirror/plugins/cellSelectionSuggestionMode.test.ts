import { describe, expect, test } from "bun:test";
import { CellSelection } from "prosemirror-tables";
import type { EditorState } from "prosemirror-state";

import {
  createHarnessState,
  HeadlessEditorView,
  parseShapeDocument,
  saveHarnessState,
} from "../../__tests__/editorHarness";
import {
  SUGGESTION_INPUT_DRIVERS,
  SUGGESTION_INPUT_KINDS,
} from "../../__tests__/suggestionInputKinds";
import { FolioDocxReviewer } from "../../ai-edits/headless";
import { fromMarkdown } from "../../markdown/fromMarkdown";
import { createDocx } from "../../docx/rezip";

const ORIGINAL_CELLS = ["First cell", "Second cell", "Third cell", "Fourth cell"];
const TABLE_MARKDOWN =
  "Before.\n\n| A | B |\n|---|---|\n| First cell | Second cell |\n| Third cell | Fourth cell |\n\nAfter.";
const focusedInputs = SUGGESTION_INPUT_KINDS.filter(
  (kind) => SUGGESTION_INPUT_DRIVERS[kind].type === "focused",
);

const bodyCells = (state: EditorState) => {
  const cells: { text: string; pos: number }[] = [];
  state.doc.descendants((node, pos) => {
    if (node.type.name === "tableCell" && ORIGINAL_CELLS.includes(node.textContent)) {
      cells.push({ text: node.textContent, pos });
    }
  });
  return cells;
};

const allBodyCellText = (state: EditorState) => {
  const cells: string[] = [];
  state.doc.descendants((node) => {
    if (node.type.name === "tableCell") {
      cells.push(node.textContent);
    }
  });
  return cells.slice(2);
};

const reopened = async (state: EditorState, base: ReturnType<typeof fromMarkdown>) => {
  const { bytes } = await saveHarnessState(state, base);
  return createHarnessState(await parseShapeDocument(bytes), "editing");
};

describe("suggestion mode cell selection deletion", () => {
  test.each(focusedInputs)(
    "%s tracks every cell in a 2×2 selection and resolves like editing",
    async (kind) => {
      if (kind !== "dragCellDelete") {
        throw new Error(`No focused input runner for ${kind}`);
      }
      const base = await parseShapeDocument(
        new Uint8Array(await createDocx(fromMarkdown(TABLE_MARKDOWN))),
      );
      const editing = new HeadlessEditorView(createHarnessState(base, "editing"));
      const suggesting = new HeadlessEditorView(createHarnessState(base, "suggesting"));

      for (const view of [editing, suggesting]) {
        const cells = bodyCells(view.state);
        expect(cells.map(({ text }) => text)).toEqual(ORIGINAL_CELLS);
        const selection = CellSelection.create(view.state.doc, cells[0]!.pos, cells[3]!.pos);
        expect(selection.ranges).toHaveLength(4);
        view.dispatch(view.state.tr.setSelection(selection));
        expect(view.pressKey("Delete")).toBe(true);
      }

      const edited = await reopened(editing.state, base);
      expect(allBodyCellText(edited)).toEqual(["", "", "", ""]);

      const suggested = await reopened(suggesting.state, base);
      expect(allBodyCellText(suggested)).toEqual(ORIGINAL_CELLS);
      for (const text of ORIGINAL_CELLS) {
        const marks: string[] = [];
        suggested.doc.descendants((node) => {
          if (node.isText && node.text === text) {
            marks.push(...node.marks.map((mark) => mark.type.name));
          }
        });
        expect(marks).toContain("deletion");
      }

      const { bytes } = await saveHarnessState(suggesting.state, base);
      const accepting = await FolioDocxReviewer.fromBuffer(bytes.slice().buffer);
      const rejecting = await FolioDocxReviewer.fromBuffer(bytes.slice().buffer);
      accepting.acceptAll();
      rejecting.rejectAll();
      const accepted = createHarnessState(
        await parseShapeDocument(new Uint8Array(await accepting.toBuffer())),
        "editing",
      );
      const rejected = createHarnessState(
        await parseShapeDocument(new Uint8Array(await rejecting.toBuffer())),
        "editing",
      );
      expect(allBodyCellText(accepted)).toEqual(allBodyCellText(edited));
      expect(allBodyCellText(rejected)).toEqual(ORIGINAL_CELLS);
    },
  );
});
