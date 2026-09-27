import { describe, expect, test } from "bun:test";
import { EditorState, TextSelection } from "prosemirror-state";

import { parseDocx } from "../../../docx/parser";
import { createDocx } from "../../../docx/rezip";
import { createEmptyDocument } from "../../../utils/createDocument";
import { fromProseDoc } from "../../conversion/fromProseDoc";
import { toProseDoc } from "../../conversion/toProseDoc";
import { createSuggestionModePlugin } from "../../plugins/suggestionMode";
import { schema } from "../../schema";
import { FootnoteRefExtension } from "./FootnoteRefExtension";
import { expandNoteReferenceDeletionRange } from "./noteReferenceDeletion";

const makeState = (id: number, noteType: "footnote" | "endnote" = "footnote") => {
  const mark = schema.mark("footnoteRef", { id: String(id), noteType });
  const doc = schema.node("doc", null, [
    schema.node("paragraph", null, [
      schema.text("Before "),
      schema.text(String(id), [mark]),
      schema.text(" after."),
    ]),
  ]);
  const start = 1 + "Before ".length;
  return { doc, start, end: start + String(id).length };
};

const press = (state: EditorState, key: "Backspace" | "Delete") => {
  const command = FootnoteRefExtension().onSchemaReady({ schema }).keyboardShortcuts?.[key];
  if (!command) {
    throw new Error(`missing ${key} shortcut`);
  }
  let result = state;
  const handled = command(state, (tr) => {
    result = result.apply(tr);
  });
  return { handled, state: result };
};

describe("note reference deletion", () => {
  for (const id of [1, 10, 100]) {
    for (const key of ["Backspace", "Delete"] as const) {
      test(`${key} removes footnote ${id} as one unit`, () => {
        const { doc, start, end } = makeState(id);
        const pos = key === "Backspace" ? end : start;
        const state = EditorState.create({
          schema,
          doc,
          selection: TextSelection.create(doc, pos),
        });

        const result = press(state, key);

        expect(result.handled).toBe(true);
        expect(result.state.doc.textContent).toBe("Before  after.");
        expect(result.state.doc.firstChild?.childCount).toBe(1);
      });
    }
  }

  test("Delete expands a partial selection across an endnote reference", () => {
    const { doc, start, end } = makeState(100, "endnote");
    const state = EditorState.create({
      schema,
      doc,
      selection: TextSelection.create(doc, start + 1, end),
    });

    expect(expandNoteReferenceDeletionRange(doc, start + 1, end)).toEqual({ from: start, to: end });
    expect(press(state, "Delete").state.doc.textContent).toBe("Before  after.");
  });

  test("a reference split by formatting still deletes as one unit", () => {
    const mark = schema.mark("footnoteRef", { id: "100", noteType: "footnote" });
    const bold = schema.mark("bold");
    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, [
        schema.text("Before "),
        schema.text("1", [mark]),
        schema.text("00", [mark, bold]),
        schema.text(" after."),
      ]),
    ]);
    const end = 1 + "Before 100".length;
    const state = EditorState.create({ schema, doc, selection: TextSelection.create(doc, end) });

    expect(press(state, "Backspace").state.doc.textContent).toBe("Before  after.");
  });

  for (const id of [10, 100]) {
    test(`deleted footnote ${id} stays absent after DOCX save and reopen`, async () => {
      const source = createEmptyDocument({ initialText: "Before  after." });
      source.package.document.content = [
        {
          type: "paragraph",
          content: [
            { type: "run", content: [{ type: "text", text: "Before " }] },
            { type: "run", content: [{ type: "footnoteRef", id }] },
            { type: "run", content: [{ type: "text", text: " after." }] },
          ],
        },
      ];
      source.package.footnotes = [
        {
          type: "footnote",
          id,
          content: [
            {
              type: "paragraph",
              content: [{ type: "run", content: [{ type: "text", text: "Note body" }] }],
            },
          ],
        },
      ];
      const doc = toProseDoc(source);
      const pos = 1 + `Before ${id}`.length;
      const state = EditorState.create({ schema, doc, selection: TextSelection.create(doc, pos) });
      const edited = fromProseDoc(press(state, "Backspace").state.doc, source);
      const saved = await createDocx(edited);
      const reopened = await parseDocx(saved, { preloadFonts: false });
      const reopenedDoc = toProseDoc(reopened);
      let hasActiveReference = false;
      reopenedDoc.descendants((node) => {
        if (node.marks.some((mark) => mark.type.name === "footnoteRef")) {
          hasActiveReference = true;
        }
      });

      expect(reopenedDoc.textContent).toBe("Before  after.");
      expect(hasActiveReference).toBe(false);
    });
  }

  test("editing shortcut yields to suggesting mode", () => {
    const { doc, end } = makeState(10);
    const state = EditorState.create({
      schema,
      doc,
      selection: TextSelection.create(doc, end),
      plugins: [createSuggestionModePlugin(true, "Reviewer")],
    });

    const result = press(state, "Backspace");
    expect(result.handled).toBe(false);
    expect(result.state.doc.eq(doc)).toBe(true);
  });
});
