import { describe, expect, test } from "bun:test";
import { Fragment, Slice } from "prosemirror-model";
import { TextSelection } from "prosemirror-state";
import { CellSelection } from "prosemirror-tables";

import { shapeArrayBuffer } from "../../__tests__/documentShapes";
import {
  createHarnessState,
  HeadlessEditorView,
  parseShapeDocument,
  saveHarnessState,
} from "../../__tests__/editorHarness";
import { FolioDocxReviewer } from "../../ai-edits/headless";
import { createCellDragTracker, findCellPosFromPmPos } from "../cellDragSelection";

const project = (reviewer: FolioDocxReviewer) =>
  reviewer.snapshot().blocks.map(({ kind, text, table }) => ({ kind, text, table }));

const tableSlice = (view: HeadlessEditorView) => {
  const { schema } = view.state;
  const cell = (text: string) =>
    schema.node("tableCell", null, schema.node("paragraph", null, schema.text(text)));
  return new Slice(
    Fragment.from(
      schema.node("table", null, [schema.node("tableRow", null, [cell("A"), cell("B")])]),
    ),
    0,
    0,
  );
};

const outsideText = (view: HeadlessEditorView) => {
  const paragraphs: string[] = [];
  view.state.doc.forEach((node) => {
    if (node.type.name === "paragraph") paragraphs.push(node.textContent);
  });
  return paragraphs.join("");
};

const cases = [0, 8, 17].flatMap((offset) =>
  (["enter", "pasteMultiBlock"] as const).flatMap((prefix) =>
    (["delete", "pasteTable"] as const).map((operation) => ({ offset, prefix, operation })),
  ),
);

describe("table drag input traces", () => {
  test.each(cases)(
    "$prefix at offset $offset then dragging cells and $operation preserves surrounding paragraphs",
    async ({ offset, prefix, operation }) => {
      const source = await shapeArrayBuffer("tables");
      const base = await parseShapeDocument(new Uint8Array(source));
      const baseline = await FolioDocxReviewer.fromBuffer(source);
      const editing = new HeadlessEditorView(createHarnessState(base, "editing"));
      const suggesting = new HeadlessEditorView(createHarnessState(base, "suggesting"));

      for (const view of [editing, suggesting]) {
        view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, offset + 1)));
        switch (prefix) {
          case "enter":
            expect(view.pressKey("Enter")).toBe(true);
            break;
          case "pasteMultiBlock": {
            const { schema } = view.state;
            view.paste(
              new Slice(
                Fragment.from([
                  schema.node("paragraph", null, schema.text("Title")),
                  schema.node("paragraph", null, schema.text("Body")),
                ]),
                1,
                1,
              ),
            );
            break;
          }
        }
        const beforeDrag = outsideText(view);
        const cells: number[] = [];
        view.state.doc.descendants((node, pos) => {
          if (node.type.name === "tableCell" && cells.length < 2) cells.push(pos);
        });
        const first = cells.at(0);
        const second = cells.at(1);
        if (first === undefined || second === undefined)
          throw new Error("Missing table drag cells");
        const tracker = createCellDragTracker();
        tracker.begin(findCellPosFromPmPos(view as never, first + 2));
        expect(tracker.update(view as never, second + 2, 0)).toBe(true);
        expect(view.state.selection).toBeInstanceOf(CellSelection);
        if (!(view.state.selection instanceof CellSelection))
          throw new Error("Missing dragged cell selection");
        expect(view.state.selection.$anchorCell.pos).toBe(first);
        expect(view.state.selection.$headCell.pos).toBe(second);
        tracker.end();
        switch (operation) {
          case "delete":
            expect(view.pressKey("Delete")).toBe(true);
            break;
          case "pasteTable":
            view.paste(tableSlice(view));
            break;
        }
        expect(outsideText(view)).toBe(beforeDrag);
      }

      const { bytes: editedBytes } = await saveHarnessState(editing.state, base);
      const edited = await FolioDocxReviewer.fromBuffer(editedBytes.slice().buffer);
      const { bytes: suggestedBytes } = await saveHarnessState(suggesting.state, base);
      const accepted = await FolioDocxReviewer.fromBuffer(suggestedBytes.slice().buffer);
      const rejected = await FolioDocxReviewer.fromBuffer(suggestedBytes.slice().buffer);
      accepted.acceptAll();
      rejected.rejectAll();
      expect(project(accepted)).toEqual(project(edited));
      expect(project(rejected)).toEqual(project(baseline));
    },
  );
});
