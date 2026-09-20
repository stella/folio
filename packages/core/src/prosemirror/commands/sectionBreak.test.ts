import { describe, test, expect } from "bun:test";
import { EditorState, TextSelection, type Transaction } from "prosemirror-state";
import { schema } from "../schema";
import { BaseKeymapExtension } from "../extensions/features/BaseKeymapExtension";
import { SECTION_BREAK_TYPES, sectionBreakTypeOf, sectionPropertiesOf } from "../sectionCarrier";
import {
  insertSectionBreakNextPage,
  insertSectionBreakContinuous,
  removeSectionBreakAtSelection,
  setSectionBreakType,
} from "./sectionBreak";

/**
 * Build an editor with a single paragraph of `text` and the cursor placed
 * `cursorOffset` characters into it (paragraph content starts at doc pos 1).
 */
function setup(text: string, cursorOffset: number) {
  const doc = schema.node("doc", { defaultTabStopTwips: null, watermark: null }, [
    schema.node("paragraph", {}, text ? [schema.text(text)] : []),
  ]);
  return editorOn(doc, 1 + cursorOffset);
}

function editorOn(doc: ReturnType<typeof schema.node>, cursorPos: number) {
  let state = EditorState.create({ doc });
  state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, cursorPos)));
  const dispatch = (tr: Transaction) => {
    state = state.apply(tr);
  };
  return {
    dispatch,
    get state() {
      return state;
    },
  };
}

const breakTypeAt = (state: EditorState, index: number) =>
  sectionBreakTypeOf(sectionPropertiesOf(state.doc.child(index)));

const backspace = () => {
  const bound = BaseKeymapExtension().onSchemaReady({ schema }).keyboardShortcuts?.["Backspace"];
  if (!bound) {
    throw new Error("the base keymap binds no Backspace");
  }
  return bound;
};

describe("insertSectionBreak commands", () => {
  test("mid-text split: first paragraph ends the section, cursor lands in the second", () => {
    const ed = setup("HelloWorld", 5);
    insertSectionBreakNextPage(ed.state, ed.dispatch);

    const { doc, selection } = ed.state;
    expect(doc.childCount).toBe(2);
    expect(breakTypeAt(ed.state, 0)).toBe("nextPage");
    expect(doc.child(0).textContent).toBe("Hello");
    // The new section's first paragraph does not itself carry a break.
    expect(sectionPropertiesOf(doc.child(1))).toBeNull();
    expect(doc.child(1).textContent).toBe("World");
    // Cursor is inside the second paragraph (the start of the new section).
    expect(selection.$from.parent.textContent).toBe("World");
  });

  test("continuous variant marks the section end as continuous", () => {
    const ed = setup("HelloWorld", 5);
    insertSectionBreakContinuous(ed.state, ed.dispatch);

    expect(breakTypeAt(ed.state, 0)).toBe("continuous");
    expect(sectionPropertiesOf(ed.state.doc.child(1))).toBeNull();
  });

  test("end-of-paragraph: current paragraph ends the section, cursor in a new empty paragraph", () => {
    const ed = setup("Hello", 5);
    insertSectionBreakNextPage(ed.state, ed.dispatch);

    const { doc, selection } = ed.state;
    expect(doc.childCount).toBe(2);
    expect(breakTypeAt(ed.state, 0)).toBe("nextPage");
    expect(doc.child(0).textContent).toBe("Hello");
    expect(doc.child(1).textContent).toBe("");
    expect(selection.$from.parent.textContent).toBe("");
  });

  test("start-of-paragraph: an empty section-ending paragraph precedes the content", () => {
    const ed = setup("Hello", 0);
    insertSectionBreakNextPage(ed.state, ed.dispatch);

    const { doc, selection } = ed.state;
    expect(doc.childCount).toBe(2);
    expect(breakTypeAt(ed.state, 0)).toBe("nextPage");
    expect(doc.child(0).textContent).toBe("");
    expect(doc.child(1).textContent).toBe("Hello");
    expect(selection.$from.parent.textContent).toBe("Hello");
  });

  test("refuses to act when the cursor is inside a table cell", () => {
    // A w:sectPr nested in a w:tc is invalid OOXML, so the command must no-op
    // (return false) rather than mark a cell paragraph as a section end.
    const cellPara = schema.node("paragraph", {}, [schema.text("in cell")]);
    const cell = schema.node("tableCell", {}, [cellPara]);
    const table = schema.node("table", {}, [schema.node("tableRow", {}, [cell])]);
    const doc = schema.node("doc", { defaultTabStopTwips: null, watermark: null }, [
      schema.node("paragraph", {}, [schema.text("body")]),
      table,
    ]);
    let state = EditorState.create({ doc });
    // Place the cursor inside the cell paragraph.
    let cursor = 0;
    doc.descendants((node, pos) => {
      if (node.isTextblock && node.textContent === "in cell") cursor = pos + 1;
      return true;
    });
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, cursor)));
    const before = state;
    const ran = insertSectionBreakNextPage(state, (tr) => {
      state = state.apply(tr);
    });
    expect(ran).toBe(false);
    expect(state).toBe(before);
  });
});

/**
 * One record carries a section break, so the commands that author one act on
 * that record rather than on a type beside it.
 *
 * Every defect below was reachable while the type was a second attr: the type
 * command could only state a type the from-leg then minted a *fresh* record
 * from, throwing away the page size, margins and header references of a parsed
 * break; the removal command cleared that type and left the parsed record in
 * place, so the break it claimed to remove was still saved; and a paragraph
 * that held the type alone had no identity for a split to be read against.
 */
describe("the one carrier under the section-break commands", () => {
  const parsedBreak = (properties: Record<string, unknown>) =>
    schema.node("doc", { defaultTabStopTwips: null, watermark: null }, [
      schema.node("paragraph", { _sectionProperties: properties }, [schema.text("Ends here")]),
      schema.node("paragraph", {}, [schema.text("Next section")]),
    ]);

  const PARSED = {
    sectionStart: "nextPage",
    pageWidth: 11_906,
    pageHeight: 16_838,
    headerReferences: [{ type: "default", rId: "rId7" }],
  } as const;

  test.each(SECTION_BREAK_TYPES)(
    "changing a parsed break to %s keeps everything the type never carried",
    (breakType) => {
      const ed = editorOn(parsedBreak({ ...PARSED }), 1);

      expect(setSectionBreakType(breakType)(ed.state, ed.dispatch)).toBe(true);

      expect(sectionPropertiesOf(ed.state.doc.child(0))).toEqual({
        ...PARSED,
        sectionStart: breakType,
      });
    },
  );

  test("removing a parsed break removes the record, not a type beside it", () => {
    const ed = editorOn(parsedBreak({ ...PARSED }), 1);

    expect(removeSectionBreakAtSelection(ed.state, ed.dispatch)).toBe(true);

    expect(sectionPropertiesOf(ed.state.doc.child(0))).toBeNull();
    // Word's rule: the paragraphs the removed break governed join the section
    // that follows, and that section's own `w:sectPr` governs them. Nothing is
    // copied onto them, because the following record already states the whole
    // of the section it heads.
    expect(sectionPropertiesOf(ed.state.doc.child(1))).toBeNull();
  });

  test("removing a break where there is none refuses instead of dispatching", () => {
    const ed = setup("No section here", 3);

    expect(removeSectionBreakAtSelection(ed.state, ed.dispatch)).toBe(false);
  });

  test("insert then remove leaves no section behind", () => {
    const ed = setup("HelloWorld", 5);
    insertSectionBreakNextPage(ed.state, ed.dispatch);

    // The cursor sits in the new section's first paragraph; select the break's
    // own paragraph to remove it.
    const back = editorOn(ed.state.doc, 1);
    expect(removeSectionBreakAtSelection(back.state, back.dispatch)).toBe(true);

    back.state.doc.descendants((node) => {
      expect(sectionPropertiesOf(node)).toBeNull();
      return true;
    });
  });

  test("two break-less paragraphs in one selection author two sections", () => {
    const doc = schema.node("doc", { defaultTabStopTwips: null, watermark: null }, [
      schema.node("paragraph", {}, [schema.text("one")]),
      schema.node("paragraph", {}, [schema.text("two")]),
    ]);
    let state = EditorState.create({ doc });
    state = state.apply(
      state.tr.setSelection(TextSelection.create(state.doc, 1, state.doc.content.size - 1)),
    );
    const ed = {
      get state() {
        return state;
      },
      dispatch: (tr: Transaction) => {
        state = state.apply(tr);
      },
    };

    expect(setSectionBreakType("continuous")(ed.state, ed.dispatch)).toBe(true);

    const first = sectionPropertiesOf(ed.state.doc.child(0));
    const second = sectionPropertiesOf(ed.state.doc.child(1));
    expect(first).toEqual({ sectionStart: "continuous" });
    // Two records, not one shared by both: a shared record is one section, and
    // these paragraphs each authored their own.
    expect(second).not.toBe(first);
  });

  test("a split's two halves keep one record when the type is changed on either", () => {
    const shared = { sectionStart: "nextPage", pageWidth: 11_906 };
    // What a split leaves: two nodes over one object, until the save leg picks
    // the last of them.
    const doc = schema.node("doc", { defaultTabStopTwips: null, watermark: null }, [
      schema.node("paragraph", { _sectionProperties: shared }, [schema.text("first half")]),
      schema.node("paragraph", { _sectionProperties: shared }, [schema.text("second half")]),
    ]);
    const ed = editorOn(doc, 1);

    expect(setSectionBreakType("continuous")(ed.state, ed.dispatch)).toBe(true);

    const first = sectionPropertiesOf(ed.state.doc.child(0));
    const second = sectionPropertiesOf(ed.state.doc.child(1));
    expect(first).toEqual({ sectionStart: "continuous", pageWidth: 11_906 });
    // Still one record. Rewriting the selected half alone would make the halves
    // two sections where the document holds one.
    expect(second).toBe(first);
  });
});

/**
 * Backspace at the start of the paragraph *after* a break deletes the break.
 *
 * That is Word's behaviour, and ECMA-376 Part 1 §17.6.18 is why: a paragraph's
 * `w:sectPr` states the properties of the section ending at *that* paragraph's
 * mark, so the join that consumes the mark takes the section with it and the
 * content falls to the following section. ProseMirror's `join` keeps the first
 * node's attrs, which left the break on a paragraph whose mark was gone.
 */
describe("Backspace against a section break", () => {
  const twoSections = () =>
    schema.node("doc", { defaultTabStopTwips: null, watermark: null }, [
      schema.node("paragraph", { _sectionProperties: { sectionStart: "nextPage" } }, [
        schema.text("ends the section"),
      ]),
      schema.node("paragraph", {}, [schema.text("after the break")]),
    ]);

  test("at the start of the paragraph after a break, the break goes", () => {
    const doc = twoSections();
    const ed = editorOn(doc, doc.child(0).nodeSize + 1);

    expect(backspace()(ed.state, ed.dispatch)).toBe(true);

    expect(ed.state.doc.childCount).toBe(1);
    expect(ed.state.doc.child(0).textContent).toBe("ends the sectionafter the break");
    expect(sectionPropertiesOf(ed.state.doc.child(0))).toBeNull();
  });

  test("at the start of the section-ending paragraph, the break stays", () => {
    // The mirror case: this join consumes the *predecessor's* mark, so the
    // merged paragraph still ends the section.
    const doc = schema.node("doc", { defaultTabStopTwips: null, watermark: null }, [
      schema.node("paragraph", {}, [schema.text("lead")]),
      schema.node("paragraph", { _sectionProperties: { sectionStart: "nextPage" } }, [
        schema.text("ends the section"),
      ]),
    ]);
    const carrier = sectionPropertiesOf(doc.child(1));
    const ed = editorOn(doc, doc.child(0).nodeSize + 1);

    expect(backspace()(ed.state, ed.dispatch)).toBe(true);

    expect(ed.state.doc.childCount).toBe(1);
    // The very object, not an equal one: the record was carried across the join
    // rather than rebuilt from a type.
    expect(sectionPropertiesOf(ed.state.doc.child(0))).toBe(carrier);
  });
});
