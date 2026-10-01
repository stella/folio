import { expect, test } from "bun:test";
import { EditorState, NodeSelection, TextSelection } from "prosemirror-state";

import { HeadlessEditorView } from "../../__tests__/editorHarness";
import { acceptAllChanges } from "../commands/comments";
import { schema } from "../schema";
import { createSuggestionModePlugin } from "./suggestionMode";

const bookmark = (type: "start" | "end", id: number) =>
  schema.node("bookmarkBoundary", {
    type,
    id,
    ...(type === "start" ? { name: `target${id}` } : {}),
  });

for (const direction of ["Backspace", "Delete"] as const) {
  for (const count of [1, 2, 3]) {
    test(`${direction} crosses ${count} bookmark starts to delete the adjacent character`, () => {
      const plugin = createSuggestionModePlugin(true, "Reviewer");
      const ids = Array.from({ length: count }, (_, index) => index + 1);
      const doc = schema.node("doc", null, [
        schema.node("paragraph", null, [
          schema.text("a"),
          ...ids.map((id) => bookmark("start", id)),
          schema.text("b"),
          ...ids.toReversed().map((id) => bookmark("end", id)),
          schema.text("c"),
        ]),
      ]);
      const caret = direction === "Backspace" ? 2 + count : 2;
      const state = EditorState.create({ doc, plugins: [plugin] });
      const view = new HeadlessEditorView(
        state.apply(state.tr.setSelection(TextSelection.create(doc, caret))),
      );

      expect(view.pressKey(direction)).toBe(true);
      const paragraph = view.state.doc.firstChild;
      expect(paragraph?.textContent).toBe("abc");
      expect(paragraph?.childCount).toBe(3 + count * 2);
      expect(
        paragraph
          ?.child(direction === "Backspace" ? 0 : count + 1)
          .marks.some(({ type }) => type.name === "deletion"),
      ).toBe(true);
      expect(paragraph?.child(1).marks.some(({ type }) => type.name === "deletion")).toBe(false);

      let accepted = view.state;
      acceptAllChanges()(accepted, (tr) => {
        accepted = accepted.apply(tr);
      });
      expect(accepted.doc.firstChild?.textContent).toBe(direction === "Backspace" ? "bc" : "ac");
      expect(
        accepted.doc.firstChild?.content.content.filter(
          ({ type }) => type.name === "bookmarkBoundary",
        ).length,
      ).toBe(count * 2);
    });

    test(`${direction} crosses ${count} bookmark pairs at a paragraph break`, () => {
      const plugin = createSuggestionModePlugin(true, "Reviewer");
      const pair = Array.from({ length: count }, (_, index) => [
        bookmark("start", index + 1),
        bookmark("end", index + 1),
      ]).flat();
      const first = schema.node(
        "paragraph",
        null,
        direction === "Delete" ? [schema.text("first"), ...pair] : [schema.text("first")],
      );
      const second = schema.node(
        "paragraph",
        null,
        direction === "Backspace" ? [...pair, schema.text("second")] : [schema.text("second")],
      );
      const doc = schema.node("doc", null, [first, second]);
      const caret =
        direction === "Backspace" ? first.nodeSize + 1 + count * 2 : first.nodeSize - 1 - count * 2;
      const state = EditorState.create({ doc, plugins: [plugin] });
      const view = new HeadlessEditorView(
        state.apply(state.tr.setSelection(TextSelection.create(doc, caret))),
      );

      expect(view.pressKey(direction)).toBe(true);
      expect(view.state.doc.firstChild?.attrs["pPrMark"]?.kind).toBe("del");
      expect(view.state.doc.textContent).toBe("firstsecond");
      let markerCount = 0;
      view.state.doc.descendants((node) => {
        if (node.type.name === "bookmarkBoundary") markerCount += 1;
      });
      expect(markerCount).toBe(count * 2);
    });
  }

  test.each([1, 2, 3])(
    `${direction} crosses %s bookmark pairs to select an adjacent table`,
    (count) => {
      const plugin = createSuggestionModePlugin(true, "Reviewer");
      const pair = Array.from({ length: count }, (_, index) => [
        bookmark("start", index + 1),
        bookmark("end", index + 1),
      ]).flat();
      const paragraph = schema.node(
        "paragraph",
        null,
        direction === "Backspace" ? [...pair, schema.text("tail")] : [schema.text("head"), ...pair],
      );
      const table = schema.node("table", null, [
        schema.node("tableRow", null, [
          schema.node("tableCell", null, [schema.node("paragraph", null, schema.text("cell"))]),
        ]),
      ]);
      const doc = schema.node(
        "doc",
        null,
        direction === "Backspace" ? [table, paragraph] : [paragraph, table],
      );
      const caret =
        direction === "Backspace"
          ? table.nodeSize + 1 + count * 2
          : paragraph.nodeSize - 1 - count * 2;
      const state = EditorState.create({ doc, plugins: [plugin] });
      const view = new HeadlessEditorView(
        state.apply(state.tr.setSelection(TextSelection.create(doc, caret))),
      );

      expect(view.pressKey(direction)).toBe(true);
      expect(view.state.selection).toBeInstanceOf(NodeSelection);
      const tablePos = direction === "Backspace" ? 0 : paragraph.nodeSize;
      expect(view.state.selection.from).toBe(tablePos);
      expect(view.state.selection.to).toBe(tablePos + table.nodeSize);
      expect(view.state.doc.eq(doc)).toBe(true);
      expect(view.state.doc.textContent).toBe(direction === "Backspace" ? "celltail" : "headcell");
      let markerCount = 0;
      view.state.doc.descendants((node) => {
        if (node.type.name === "bookmarkBoundary") markerCount += 1;
      });
      expect(markerCount).toBe(count * 2);
    },
  );
}
