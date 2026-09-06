import { describe, expect, test } from "bun:test";
import { Schema } from "prosemirror-model";
import { EditorState } from "prosemirror-state";
import type { Transaction } from "prosemirror-state";

import { acceptAllChanges, acceptChange, rejectAllChanges, rejectChange } from "./comments";

const schema = new Schema({
  nodes: {
    doc: { content: "block+" },
    paragraph: {
      content: "inline*",
      group: "block",
      attrs: {
        pPrMark: { default: null },
        sectionBreakType: { default: null },
        _sectionProperties: { default: null },
      },
    },
    text: { group: "inline", marks: "_" },
    // Zero-width anchors: they hold a position, carry no revision of their
    // own, and therefore outlive a deletion that took every word around them.
    renderedPageBreak: { inline: true, group: "inline", atom: true },
    bookmarkBoundary: { inline: true, group: "inline", atom: true },
    textBoxAnchor: { inline: true, group: "inline", atom: true },
  },
  marks: {
    insertion: {
      attrs: { revisionId: {}, author: {}, date: {} },
      excludes: "",
      toDOM: () => ["ins", 0],
    },
    deletion: {
      attrs: { revisionId: {}, author: {}, date: {} },
      excludes: "",
      toDOM: () => ["del", 0],
    },
  },
});

const dispatcher = (state: EditorState) => {
  const view = {
    state,
    dispatch(tr: Transaction) {
      view.state = view.state.apply(tr);
    },
  };
  return view;
};

const insMark = (info: { id: number; author?: string }) => ({
  kind: "ins" as const,
  info: { id: info.id, author: info.author ?? "Alice" },
});

const delMark = (info: { id: number; author?: string }) => ({
  kind: "del" as const,
  info: { id: info.id, author: info.author ?? "Alice" },
});

const twoParagraphs = (firstPPrMark: unknown) =>
  EditorState.create({
    schema,
    doc: schema.node("doc", null, [
      schema.node("paragraph", { pPrMark: firstPPrMark }, schema.text("first")),
      schema.node("paragraph", null, schema.text("second")),
    ]),
  });

describe("pPrMark accept / reject — paragraph-mark resolution", () => {
  test("accept clears pPrMark.kind = 'ins' (the paragraph break stays)", () => {
    const view = dispatcher(twoParagraphs(insMark({ id: 1 })));
    acceptAllChanges()(view.state, view.dispatch);

    expect(view.state.doc.childCount).toBe(2);
    expect(view.state.doc.child(0).attrs["pPrMark"]).toBeNull();
    expect(view.state.doc.child(0).textContent).toBe("first");
    expect(view.state.doc.child(1).textContent).toBe("second");
  });

  test("reject of pPrMark.kind = 'ins' joins this paragraph with the next", () => {
    const view = dispatcher(twoParagraphs(insMark({ id: 1 })));
    rejectAllChanges()(view.state, view.dispatch);

    expect(view.state.doc.childCount).toBe(1);
    expect(view.state.doc.child(0).textContent).toBe("firstsecond");
    expect(view.state.doc.child(0).attrs["pPrMark"]).toBeNull();
  });

  test("accept of pPrMark.kind = 'del' joins this paragraph with the next", () => {
    const view = dispatcher(twoParagraphs(delMark({ id: 1 })));
    acceptAllChanges()(view.state, view.dispatch);

    expect(view.state.doc.childCount).toBe(1);
    expect(view.state.doc.child(0).textContent).toBe("firstsecond");
    expect(view.state.doc.child(0).attrs["pPrMark"]).toBeNull();
  });

  test("reject clears pPrMark.kind = 'del' (the paragraph break stays)", () => {
    const view = dispatcher(twoParagraphs(delMark({ id: 1 })));
    rejectAllChanges()(view.state, view.dispatch);

    expect(view.state.doc.childCount).toBe(2);
    expect(view.state.doc.child(0).attrs["pPrMark"]).toBeNull();
    expect(view.state.doc.child(0).textContent).toBe("first");
    expect(view.state.doc.child(1).textContent).toBe("second");
  });

  test("range-scoped acceptChange ignores paragraphs outside the range", () => {
    const state = EditorState.create({
      schema,
      doc: schema.node("doc", null, [
        schema.node("paragraph", { pPrMark: insMark({ id: 1 }) }, schema.text("p1")),
        schema.node("paragraph", { pPrMark: insMark({ id: 2 }) }, schema.text("p2")),
        schema.node("paragraph", null, schema.text("p3")),
      ]),
    });
    const view = dispatcher(state);

    // Range covers only the first paragraph, including its closing boundary.
    acceptChange(0, 4)(view.state, view.dispatch);

    expect(view.state.doc.child(0).attrs["pPrMark"]).toBeNull();
    expect(view.state.doc.child(1).attrs["pPrMark"]).toEqual(insMark({ id: 2 }));
  });

  test("range-scoped acceptChange keeps pPrMark when only inline text is selected", () => {
    const insertion = schema.marks["insertion"]!;
    const pPrMark = insMark({ id: 1 });
    const state = EditorState.create({
      schema,
      doc: schema.node("doc", null, [
        schema.node("paragraph", { pPrMark }, [
          schema.text("a"),
          schema.text("b", [
            insertion.create({
              revisionId: 2,
              author: "Alice",
              date: "2026-05-01",
            }),
          ]),
        ]),
        schema.node("paragraph", null, schema.text("next")),
      ]),
    });
    const view = dispatcher(state);

    acceptChange(2, 3)(view.state, view.dispatch);

    let hasInsertion = false;
    view.state.doc.child(0).descendants((node) => {
      if (node.isText && node.marks.some((mark) => mark.type === insertion)) {
        hasInsertion = true;
      }
    });
    expect(hasInsertion).toBe(false);
    expect(view.state.doc.child(0).attrs["pPrMark"]).toEqual(pPrMark);
    expect(view.state.doc.childCount).toBe(2);
  });

  test("acceptAll on a doc-terminal pPrMark='del' clears a marker it cannot join", () => {
    // The paragraph keeps its words, so there is nothing to remove: only the
    // break it claims went, and there is no paragraph after it to join with.
    // The revision is resolved rather than left standing over a document that
    // no longer carries it.
    const state = EditorState.create({
      schema,
      doc: schema.node("doc", null, [
        schema.node("paragraph", null, schema.text("first")),
        schema.node("paragraph", { pPrMark: delMark({ id: 1 }) }, schema.text("last")),
      ]),
    });
    const view = dispatcher(state);
    acceptAllChanges()(view.state, view.dispatch);

    expect(view.state.doc.childCount).toBe(2);
    expect(view.state.doc.child(1).attrs["pPrMark"]).toBeNull();
  });

  test("acceptAll removes an emptied paragraph it cannot join", () => {
    // A paragraph whose words and whose break were both resolved away is not
    // there any more. Left blank it would be a line the accepted document
    // never had, which is what happens to a paragraph deleted before a table.
    const state = EditorState.create({
      schema,
      doc: schema.node("doc", null, [
        schema.node("paragraph", null, schema.text("first")),
        schema.node("paragraph", { pPrMark: delMark({ id: 1 }) }),
      ]),
    });
    const view = dispatcher(state);
    acceptAllChanges()(view.state, view.dispatch);

    expect(view.state.doc.childCount).toBe(1);
    expect(view.state.doc.child(0).textContent).toBe("first");
  });

  // A zero-width anchor holds a position and shows nothing. Counting one as
  // content left an emptied paragraph standing as a blank line around an
  // invisible node.
  for (const anchor of ["renderedPageBreak", "bookmarkBoundary", "textBoxAnchor"]) {
    test(`a ${anchor} does not keep an emptied paragraph alive`, () => {
      const deletion = schema.marks["deletion"]!;
      const revision = { revisionId: 1, author: "Alice", date: "2026-05-01" };
      const state = EditorState.create({
        schema,
        doc: schema.node("doc", null, [
          schema.node("paragraph", null, schema.text("first")),
          schema.node("paragraph", { pPrMark: delMark({ id: 2 }) }, [
            schema.node(anchor),
            schema.text("gone", [deletion.create(revision)]),
          ]),
        ]),
      });
      const view = dispatcher(state);

      acceptAllChanges()(view.state, view.dispatch);

      expect(view.state.doc.childCount).toBe(1);
      expect(view.state.doc.child(0).textContent).toBe("first");
    });
  }

  // A section break lives on a paragraph mark, so the paragraph is where the
  // section ends. Resolving the mark away must not take the section with it:
  // the dropped section's page size, margins and header and footer references
  // would go too, and a document that had two would end up with one.
  describe("a section on the resolved paragraph's mark", () => {
    const deletion = () => schema.marks["deletion"]!;
    const revision = { revisionId: 1, author: "Alice", date: "2026-05-01" };

    test("travels to the paragraph the join leaves behind", () => {
      const state = EditorState.create({
        schema,
        doc: schema.node("doc", null, [
          schema.node(
            "paragraph",
            { pPrMark: delMark({ id: 2 }), sectionBreakType: "nextPage" },
            schema.text("gone", [deletion().create(revision)]),
          ),
          schema.node("paragraph", null, schema.text("next")),
        ]),
      });
      const view = dispatcher(state);

      acceptAllChanges()(view.state, view.dispatch);

      expect(view.state.doc.childCount).toBe(1);
      expect(view.state.doc.child(0).textContent).toBe("next");
      expect(view.state.doc.child(0).attrs["sectionBreakType"]).toBe("nextPage");
    });

    test("moves back a paragraph when there is nothing to join it with", () => {
      const state = EditorState.create({
        schema,
        doc: schema.node("doc", null, [
          schema.node("paragraph", null, schema.text("first")),
          schema.node(
            "paragraph",
            { pPrMark: delMark({ id: 2 }), sectionBreakType: "nextPage" },
            schema.text("gone", [deletion().create(revision)]),
          ),
        ]),
      });
      const view = dispatcher(state);

      acceptAllChanges()(view.state, view.dispatch);

      // The section now ends one paragraph earlier; the content before it is
      // still in that section.
      expect(view.state.doc.childCount).toBe(1);
      expect(view.state.doc.child(0).textContent).toBe("first");
      expect(view.state.doc.child(0).attrs["sectionBreakType"]).toBe("nextPage");
    });

    test("keeps the paragraph when the one before it ends a section of its own", () => {
      const state = EditorState.create({
        schema,
        doc: schema.node("doc", null, [
          schema.node("paragraph", { sectionBreakType: "continuous" }, schema.text("first")),
          schema.node(
            "paragraph",
            { pPrMark: delMark({ id: 2 }), sectionBreakType: "nextPage" },
            schema.text("gone", [deletion().create(revision)]),
          ),
        ]),
      });
      const view = dispatcher(state);

      acceptAllChanges()(view.state, view.dispatch);

      expect(view.state.doc.childCount).toBe(2);
      expect(view.state.doc.child(0).attrs["sectionBreakType"]).toBe("continuous");
      expect(view.state.doc.child(1).attrs["sectionBreakType"]).toBe("nextPage");
      expect(view.state.doc.child(1).attrs["pPrMark"]).toBeNull();
    });
  });

  test("rejectChange + inline insertion on same paragraph resolves both", () => {
    const insertion = schema.marks["insertion"]!;
    const state = EditorState.create({
      schema,
      doc: schema.node("doc", null, [
        schema.node("paragraph", { pPrMark: insMark({ id: 1 }) }, [
          schema.text("kept "),
          schema.text("inserted", [
            insertion.create({
              revisionId: 1,
              author: "Alice",
              date: "2026-05-01",
            }),
          ]),
        ]),
        schema.node("paragraph", null, schema.text("next")),
      ]),
    });
    const view = dispatcher(state);

    rejectChange(0, view.state.doc.content.size)(view.state, view.dispatch);

    expect(view.state.doc.childCount).toBe(1);
    expect(view.state.doc.child(0).textContent).toBe("kept next");
    expect(view.state.doc.child(0).attrs["pPrMark"]).toBeNull();
  });
});
