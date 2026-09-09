import { describe, expect, test } from "bun:test";
import { NodeSelection, EditorState, TextSelection, type Transaction } from "prosemirror-state";

import { schema } from "../schema";
import { getTrackedChangesFromDoc } from "../../ai-edits/read";
import { createSuggestionModePlugin } from "../plugins/suggestionMode";
import { insertPageBreak } from "./pageBreak";

function setup(text: string, cursorOffset: number) {
  const doc = schema.node("doc", null, [
    schema.node("paragraph", null, text ? [schema.text(text)] : []),
  ]);
  let state = EditorState.create({ doc });
  state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 1 + cursorOffset)));
  const dispatch = (transaction: Transaction): void => {
    state = state.apply(transaction);
  };
  return {
    dispatch,
    get state() {
      return state;
    },
  };
}

describe("insertPageBreak", () => {
  test.each([
    ["leading", 0, ["pageBreakRun", "text"]],
    ["interior", 2, ["text", "pageBreakRun", "text"]],
    ["trailing", 4, ["text", "pageBreakRun"]],
  ])("inserts one inline carrier at a %s cursor", (_position, offset, expectedTypes) => {
    const editor = setup("Text", offset);

    expect(insertPageBreak(editor.state, editor.dispatch)).toBe(true);

    const paragraph = editor.state.doc.child(0);
    const childTypes = Array.from(
      { length: paragraph.childCount },
      (_, index) => paragraph.child(index).type.name,
    );
    expect(editor.state.doc.childCount).toBe(1);
    expect(childTypes).toEqual(expectedTypes);
    expect(editor.state.selection.empty).toBe(true);
    expect(editor.state.selection.$from.parent).toBe(paragraph);
  });

  test("inherits the authored run formatting at the insertion cursor", () => {
    const bold = schema.marks["bold"]?.create();
    if (!bold) {
      throw new Error("Expected bold mark");
    }
    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, [schema.text("Text", [bold])]),
    ]);
    let state = EditorState.create({ doc });
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 3)));

    expect(
      insertPageBreak(state, (transaction) => {
        state = state.apply(transaction);
      }),
    ).toBe(true);

    expect(state.doc.child(0).child(1).type.name).toBe("pageBreakRun");
    expect(
      state.doc
        .child(0)
        .child(1)
        .marks.map(({ type }) => type.name),
    ).toContain("bold");
  });

  test("allocates and enumerates a revision when inserted in suggestion mode", () => {
    const doc = schema.node("doc", null, [schema.node("paragraph", null, [schema.text("Text")])]);
    let state = EditorState.create({
      doc,
      plugins: [createSuggestionModePlugin(true, "Reviewer")],
    });
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 3)));

    expect(
      insertPageBreak(state, (transaction) => {
        state = state.apply(transaction);
      }),
    ).toBe(true);

    const pageBreak = state.doc.child(0).child(1);
    const insertion = pageBreak.marks.find(({ type }) => type.name === "insertion");
    expect(insertion?.attrs["revisionId"]).toBeGreaterThan(0);
    expect(insertion?.attrs["author"]).toBe("Reviewer");
    expect(getTrackedChangesFromDoc(state.doc)).toEqual([
      expect.objectContaining({
        id: insertion?.attrs["revisionId"],
        author: "Reviewer",
        type: "insertion",
        text: "",
      }),
    ]);
  });

  test("refuses a non-empty text replacement", () => {
    const editor = setup("Text", 0);
    const selected = editor.state.apply(
      editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 1, 3)),
    );

    expect(insertPageBreak(selected, editor.dispatch)).toBe(false);
    expect(editor.state.doc.textContent).toBe("Text");
  });

  test("refuses a block selection at the legacy boundary", () => {
    const doc = schema.node("doc", null, [
      schema.node("pageBreak"),
      schema.node("paragraph", null, [schema.text("After")]),
    ]);
    let state = EditorState.create({ doc });
    state = state.apply(state.tr.setSelection(NodeSelection.create(state.doc, 0)));

    expect(insertPageBreak(state)).toBe(false);
  });

  test.each(["tableCell", "tableHeader"] as const)(
    "refuses a cursor inside a %s before creating a transaction",
    (cellType) => {
      const paragraph = schema.node("paragraph", null, [schema.text("Cell")]);
      const doc = schema.node("doc", null, [
        schema.node("table", null, [
          schema.node("tableRow", null, [schema.node(cellType, null, [paragraph])]),
        ]),
      ]);
      let paragraphPosition: number | undefined;
      doc.descendants((node, position) => {
        if (node === paragraph) {
          paragraphPosition = position;
        }
      });
      if (paragraphPosition === undefined) {
        throw new Error("Expected cell paragraph");
      }
      let state = EditorState.create({ doc });
      state = state.apply(
        state.tr.setSelection(TextSelection.create(state.doc, paragraphPosition + 1)),
      );
      let dispatched = false;

      expect(
        insertPageBreak(state, () => {
          dispatched = true;
        }),
      ).toBe(false);
      expect(dispatched).toBe(false);
      expect(state.doc).toBe(doc);
    },
  );

  test("refuses a cursor inside a text box before creating a transaction", () => {
    const paragraph = schema.node("paragraph", null, [schema.text("Box")]);
    const doc = schema.node("doc", null, [schema.node("textBox", { width: 120 }, [paragraph])]);
    let paragraphPosition: number | undefined;
    doc.descendants((node, position) => {
      if (node === paragraph) {
        paragraphPosition = position;
      }
    });
    if (paragraphPosition === undefined) {
      throw new Error("Expected text-box paragraph");
    }
    let state = EditorState.create({ doc });
    state = state.apply(
      state.tr.setSelection(TextSelection.create(state.doc, paragraphPosition + 1)),
    );
    let dispatched = false;

    expect(
      insertPageBreak(state, () => {
        dispatched = true;
      }),
    ).toBe(false);
    expect(dispatched).toBe(false);
    expect(state.doc).toBe(doc);
  });

  test.each([
    ["frame", { _originalFormatting: { frame: { width: 720 } } }],
    ["outline", { outlineLevel: 0 }],
    ["borders", { borders: { bottom: { style: "single", size: 8 } } }],
  ] as const)("refuses a cursor in a paragraph with %s ownership", (_, attrs) => {
    const doc = schema.node("doc", null, [schema.node("paragraph", attrs, [schema.text("Text")])]);
    let state = EditorState.create({ doc });
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 2)));
    let dispatched = false;

    expect(
      insertPageBreak(state, () => {
        dispatched = true;
      }),
    ).toBe(false);
    expect(dispatched).toBe(false);
    expect(state.doc).toBe(doc);
  });

  test.each(["drop", "margin"] as const)(
    "keeps page-break insertion available beside a %s-cap frame",
    (dropCap) => {
      const doc = schema.node("doc", null, [
        schema.node("paragraph", { _originalFormatting: { frame: { dropCap } } }, [
          schema.text("Text"),
        ]),
      ]);
      let state = EditorState.create({ doc });
      state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 2)));

      expect(
        insertPageBreak(state, (transaction) => {
          state = state.apply(transaction);
        }),
      ).toBe(true);
      expect(state.doc.child(0).child(1).type.name).toBe("pageBreakRun");
    },
  );

  test("refuses a paragraph containing a text-box anchor before creating a transaction", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, [
        schema.text("AB"),
        schema.node("textBoxAnchor", { anchorId: "paragraph:0" }),
      ]),
    ]);
    let state = EditorState.create({ doc });
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 2)));
    let dispatched = false;

    expect(
      insertPageBreak(state, () => {
        dispatched = true;
      }),
    ).toBe(false);
    expect(dispatched).toBe(false);
    expect(state.doc).toBe(doc);
  });
});
