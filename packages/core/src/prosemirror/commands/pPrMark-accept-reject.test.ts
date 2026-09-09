import { describe, expect, test } from "bun:test";
import { Schema } from "prosemirror-model";
import { EditorState } from "prosemirror-state";
import type { Transaction } from "prosemirror-state";

import {
  getTrackedSectionEndpointRemoval,
  ParagraphChangeTrackerExtension,
} from "../extensions/features/ParagraphChangeTrackerExtension";
import {
  acceptAIEditRevision,
  acceptAllChanges,
  acceptChange,
  rejectAllChanges,
  rejectChange,
} from "./comments";

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
    // A block a paragraph mark cannot join with, and which cannot end a
    // container: what a table is to the paragraph before it.
    table: { content: "block+", group: "block" },
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

const trackerRuntime = ParagraphChangeTrackerExtension().onSchemaReady({ schema });
const trackerPlugin = trackerRuntime.plugins?.at(0);
if (!trackerPlugin) {
  throw new Error("Expected paragraph change tracker plugin");
}

const trackedState = (doc: EditorState["doc"]) =>
  EditorState.create({ schema, doc, plugins: [trackerPlugin] });

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
        schema.node("paragraph", { pPrMark: delMark({ id: 1 }) }),
        schema.node("table", null, [schema.node("paragraph", null, schema.text("cell"))]),
        schema.node("paragraph", null, schema.text("last")),
      ]),
    });
    const view = dispatcher(state);
    acceptAllChanges()(view.state, view.dispatch);

    expect(view.state.doc.childCount).toBe(2);
    expect(view.state.doc.child(0).type.name).toBe("table");
  });

  test("acceptAll keeps an emptied paragraph that ENDS its container", () => {
    // Nothing follows it, so the break it claims went cannot be closed over
    // anything and the container would be left without the paragraph it has to
    // end with. The revision is resolved; the paragraph stays, blank.
    const state = EditorState.create({
      schema,
      doc: schema.node("doc", null, [
        schema.node("paragraph", null, schema.text("first")),
        schema.node("paragraph", { pPrMark: delMark({ id: 1 }) }),
      ]),
    });
    const view = dispatcher(state);
    acceptAllChanges()(view.state, view.dispatch);

    expect(view.state.doc.childCount).toBe(2);
    expect(view.state.doc.child(1).textContent).toBe("");
    expect(view.state.doc.child(1).attrs["pPrMark"]).toBeNull();
  });

  test("rejectAll still removes an emptied paragraph an insertion ADDED at the end", () => {
    // The mirror case: the break was added, so the paragraph it ends was not
    // there before it and taking the addition back leaves the container ending
    // where it did.
    const insertion = schema.marks["insertion"]!;
    const state = EditorState.create({
      schema,
      doc: schema.node("doc", null, [
        schema.node("paragraph", null, schema.text("first")),
        schema.node(
          "paragraph",
          { pPrMark: insMark({ id: 1 }) },
          schema.text("added", [
            insertion.create({ revisionId: 1, author: "Alice", date: "2026-05-01" }),
          ]),
        ),
      ]),
    });
    const view = dispatcher(state);
    rejectAllChanges()(view.state, view.dispatch);

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
          schema.node("paragraph", { pPrMark: delMark({ id: 2 }) }, [
            schema.node(anchor),
            schema.text("gone", [deletion.create(revision)]),
          ]),
          schema.node("table", null, [schema.node("paragraph", null, schema.text("cell"))]),
          schema.node("paragraph", null, schema.text("last")),
        ]),
      });
      const view = dispatcher(state);

      acceptAllChanges()(view.state, view.dispatch);

      expect(view.state.doc.childCount).toBe(2);
      expect(view.state.doc.child(0).type.name).toBe("table");
    });
  }

  // A section break lives on a paragraph mark, so the paragraph is where the
  // section ends. Resolving that mark away removes the endpoint with it; the
  // following paragraph's mark owns the section that survives the join.
  describe("a section on the resolved paragraph's mark", () => {
    const deletion = () => schema.marks["deletion"]!;
    const revision = { revisionId: 1, author: "Alice", date: "2026-05-01" };

    test("is removed when an ordinary following paragraph survives", () => {
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
      expect(view.state.doc.child(0).attrs["sectionBreakType"]).toBeNull();
    });

    test("removes only the source endpoint when the next paragraph owns one", () => {
      const state = EditorState.create({
        schema,
        doc: schema.node("doc", null, [
          schema.node(
            "paragraph",
            { pPrMark: delMark({ id: 2 }), sectionBreakType: "nextPage" },
            schema.text("first section"),
          ),
          schema.node(
            "paragraph",
            { sectionBreakType: "continuous" },
            schema.text("second section"),
          ),
        ]),
      });
      const view = dispatcher(state);

      acceptAllChanges()(view.state, view.dispatch);

      expect(view.state.doc.childCount).toBe(1);
      expect(view.state.doc.child(0).textContent).toBe("first sectionsecond section");
      expect(view.state.doc.child(0).attrs["sectionBreakType"]).toBe("continuous");
      expect(view.state.doc.child(0).attrs["pPrMark"]).toBeNull();
    });

    test("keeps the next paragraph's section when surviving content is joined", () => {
      const state = EditorState.create({
        schema,
        doc: schema.node("doc", null, [
          schema.node("paragraph", { pPrMark: delMark({ id: 2 }) }, schema.text("first")),
          schema.node(
            "paragraph",
            {
              pPrMark: insMark({ id: 3 }),
              sectionBreakType: "continuous",
              _sectionProperties: { columns: 2 },
            },
            schema.text("second"),
          ),
        ]),
      });
      const view = dispatcher(state);

      acceptChange(0, view.state.doc.child(0).nodeSize)(view.state, view.dispatch);

      expect(view.state.doc.childCount).toBe(1);
      expect(view.state.doc.child(0).textContent).toBe("firstsecond");
      expect(view.state.doc.child(0).attrs["sectionBreakType"]).toBe("continuous");
      expect(view.state.doc.child(0).attrs["_sectionProperties"]).toEqual({ columns: 2 });
      expect(view.state.doc.child(0).attrs["pPrMark"]).toEqual(insMark({ id: 3 }));
    });

    test("keeps the next paragraph's revision and section after emptied content is joined", () => {
      const deletionMark = deletion().create(revision);
      const state = EditorState.create({
        schema,
        doc: schema.node("doc", null, [
          schema.node(
            "paragraph",
            { pPrMark: delMark({ id: 2 }) },
            schema.text("gone", [deletionMark]),
          ),
          schema.node(
            "paragraph",
            {
              pPrMark: insMark({ id: 3 }),
              sectionBreakType: "continuous",
              _sectionProperties: { columns: 2 },
            },
            schema.text("second"),
          ),
        ]),
      });
      const view = dispatcher(state);

      acceptChange(0, view.state.doc.child(0).nodeSize)(view.state, view.dispatch);

      expect(view.state.doc.childCount).toBe(1);
      expect(view.state.doc.child(0).textContent).toBe("second");
      expect(view.state.doc.child(0).attrs["sectionBreakType"]).toBe("continuous");
      expect(view.state.doc.child(0).attrs["_sectionProperties"]).toEqual({ columns: 2 });
      expect(view.state.doc.child(0).attrs["pPrMark"]).toEqual(insMark({ id: 3 }));
    });

    test("is removed with an emptied paragraph before a nonjoinable sibling", () => {
      const state = EditorState.create({
        schema,
        doc: schema.node("doc", null, [
          schema.node("paragraph", null, schema.text("first")),
          schema.node(
            "paragraph",
            { pPrMark: delMark({ id: 2 }), sectionBreakType: "nextPage" },
            schema.text("gone", [deletion().create(revision)]),
          ),
          schema.node("table", null, [schema.node("paragraph", null, schema.text("cell"))]),
          schema.node("paragraph", null, schema.text("last")),
        ]),
      });
      const view = dispatcher(state);

      acceptAllChanges()(view.state, view.dispatch);

      expect(view.state.doc.childCount).toBe(3);
      expect(view.state.doc.child(0).textContent).toBe("first");
      expect(view.state.doc.child(0).attrs["sectionBreakType"]).toBeNull();
    });

    test("leaves a preceding section endpoint unchanged before a nonjoinable sibling", () => {
      const state = EditorState.create({
        schema,
        doc: schema.node("doc", null, [
          schema.node("paragraph", { sectionBreakType: "continuous" }, schema.text("first")),
          schema.node(
            "paragraph",
            { pPrMark: delMark({ id: 2 }), sectionBreakType: "nextPage" },
            schema.text("gone", [deletion().create(revision)]),
          ),
          schema.node("table", null, [schema.node("paragraph", null, schema.text("cell"))]),
          schema.node("paragraph", null, schema.text("last")),
        ]),
      });
      const view = dispatcher(state);

      acceptAllChanges()(view.state, view.dispatch);

      expect(view.state.doc.childCount).toBe(3);
      expect(view.state.doc.child(0).attrs["sectionBreakType"]).toBe("continuous");
      expect(view.state.doc.child(1).type.name).toBe("table");
    });

    test("reject keeps a deleted endpoint on its original paragraph", () => {
      const state = EditorState.create({
        schema,
        doc: schema.node("doc", null, [
          schema.node(
            "paragraph",
            {
              pPrMark: delMark({ id: 2 }),
              sectionBreakType: "nextPage",
              _sectionProperties: { columns: 2 },
            },
            schema.text("source section"),
          ),
          schema.node("paragraph", null, schema.text("following section")),
        ]),
      });
      const view = dispatcher(state);

      rejectAllChanges()(view.state, view.dispatch);

      expect(view.state.doc.childCount).toBe(2);
      expect(view.state.doc.child(0).attrs["pPrMark"]).toBeNull();
      expect(view.state.doc.child(0).attrs["sectionBreakType"]).toBe("nextPage");
      expect(view.state.doc.child(0).attrs["_sectionProperties"]).toEqual({ columns: 2 });
    });

    test("reject removes an inserted endpoint and retains the following paragraph", () => {
      const state = EditorState.create({
        schema,
        doc: schema.node("doc", null, [
          schema.node(
            "paragraph",
            {
              pPrMark: insMark({ id: 2 }),
              sectionBreakType: "nextPage",
              _sectionProperties: { columns: 2 },
            },
            schema.text("inserted section"),
          ),
          schema.node("paragraph", null, schema.text("following section")),
        ]),
      });
      const view = dispatcher(state);

      rejectAllChanges()(view.state, view.dispatch);

      expect(view.state.doc.childCount).toBe(1);
      expect(view.state.doc.child(0).textContent).toBe("inserted sectionfollowing section");
      expect(view.state.doc.child(0).attrs["sectionBreakType"]).toBeNull();
    });

    test("authorizes only the exact endpoint-removal direction", () => {
      const sectionParagraph = () =>
        schema.node(
          "paragraph",
          {
            pPrMark: delMark({ id: 2 }),
            sectionBreakType: "nextPage",
            _sectionProperties: {
              columns: 2,
              headerReferences: [{ type: "first", rId: "rId7" }],
            },
          },
          schema.text("source section"),
        );
      const followingParagraph = () =>
        schema.node("paragraph", null, schema.text("following section"));
      const source = schema.node("doc", null, [sectionParagraph(), followingParagraph()]);
      const accepting = dispatcher(trackedState(source));
      const rejecting = dispatcher(trackedState(source));

      acceptAllChanges()(accepting.state, accepting.dispatch);
      rejectAllChanges()(rejecting.state, rejecting.dispatch);

      const authorization = getTrackedSectionEndpointRemoval(accepting.state);
      expect(authorization).toMatchObject({
        type: "tracked-section-endpoint-removal",
        sourceParagraphEndpointCount: 1,
        expectedParagraphEndpointCount: 0,
        removedReferences: [{ part: "header", type: "first", relationshipId: "rId7" }],
      });
      expect(authorization?.sourceEndpointFingerprint).not.toBe(
        authorization?.expectedEndpointFingerprint,
      );
      expect(getTrackedSectionEndpointRemoval(rejecting.state)).toBeNull();

      const nonSection = dispatcher(trackedState(twoParagraphs(delMark({ id: 2 })).doc));
      acceptAllChanges()(nonSection.state, nonSection.dispatch);
      expect(getTrackedSectionEndpointRemoval(nonSection.state)).toBeNull();
    });

    test("counts only section-owning marks in a mixed accept-all batch", () => {
      const state = trackedState(
        schema.node("doc", null, [
          schema.node(
            "paragraph",
            {
              pPrMark: delMark({ id: 2 }),
              sectionBreakType: "nextPage",
              _sectionProperties: { columns: 2 },
            },
            schema.text("section paragraph"),
          ),
          schema.node(
            "paragraph",
            { pPrMark: delMark({ id: 3 }) },
            schema.text("ordinary paragraph"),
          ),
          schema.node("paragraph", null, schema.text("following paragraph")),
        ]),
      );
      const view = dispatcher(state);

      acceptAllChanges()(view.state, view.dispatch);

      expect(getTrackedSectionEndpointRemoval(view.state)).toMatchObject({
        sourceParagraphEndpointCount: 1,
        expectedParagraphEndpointCount: 0,
      });
    });

    test("invalidates an authorization after a later untracked endpoint removal", () => {
      const state = trackedState(
        schema.node("doc", null, [
          schema.node(
            "paragraph",
            {
              pPrMark: delMark({ id: 2 }),
              sectionBreakType: "nextPage",
              _sectionProperties: { columns: 2 },
            },
            schema.text("removed endpoint"),
          ),
          schema.node(
            "paragraph",
            {
              sectionBreakType: "continuous",
              _sectionProperties: { columns: 3 },
            },
            schema.text("surviving endpoint"),
          ),
        ]),
      );
      const view = dispatcher(state);
      acceptAllChanges()(view.state, view.dispatch);
      expect(getTrackedSectionEndpointRemoval(view.state)).toMatchObject({
        sourceParagraphEndpointCount: 2,
        expectedParagraphEndpointCount: 1,
      });

      view.dispatch(
        view.state.tr
          .setNodeAttribute(0, "sectionBreakType", null)
          .setNodeAttribute(0, "_sectionProperties", null),
      );

      expect(getTrackedSectionEndpointRemoval(view.state)).toBeNull();
    });

    test("invalidates an authorization after a count-neutral endpoint property change", () => {
      const state = trackedState(
        schema.node("doc", null, [
          schema.node(
            "paragraph",
            {
              pPrMark: delMark({ id: 2 }),
              sectionBreakType: "nextPage",
              _sectionProperties: { columns: 2 },
            },
            schema.text("removed endpoint"),
          ),
          schema.node(
            "paragraph",
            {
              sectionBreakType: "continuous",
              _sectionProperties: { columns: 3 },
            },
            schema.text("surviving endpoint"),
          ),
        ]),
      );
      const view = dispatcher(state);
      acceptAllChanges()(view.state, view.dispatch);
      expect(getTrackedSectionEndpointRemoval(view.state)).not.toBeNull();

      view.dispatch(view.state.tr.setNodeAttribute(0, "_sectionProperties", { columns: 4 }));

      expect(getTrackedSectionEndpointRemoval(view.state)).toBeNull();
    });

    test("invalidates an authorization after a count-neutral endpoint move", () => {
      const state = trackedState(
        schema.node("doc", null, [
          schema.node(
            "paragraph",
            {
              pPrMark: delMark({ id: 2 }),
              sectionBreakType: "nextPage",
              _sectionProperties: { columns: 2 },
            },
            schema.text("removed endpoint"),
          ),
          schema.node(
            "paragraph",
            {
              sectionBreakType: "continuous",
              _sectionProperties: { columns: 3 },
            },
            schema.text("surviving endpoint"),
          ),
          schema.node("paragraph", null, schema.text("ordinary paragraph")),
        ]),
      );
      const view = dispatcher(state);
      acceptAllChanges()(view.state, view.dispatch);
      expect(getTrackedSectionEndpointRemoval(view.state)).not.toBeNull();
      const firstParagraphSize = view.state.doc.child(0).nodeSize;

      view.dispatch(
        view.state.tr
          .setNodeAttribute(0, "sectionBreakType", null)
          .setNodeAttribute(0, "_sectionProperties", null)
          .setNodeAttribute(firstParagraphSize, "sectionBreakType", "continuous")
          .setNodeAttribute(firstParagraphSize, "_sectionProperties", { columns: 3 }),
      );

      expect(getTrackedSectionEndpointRemoval(view.state)).toBeNull();
    });

    test("accumulates exact endpoint removals across scoped resolutions", () => {
      const state = trackedState(
        schema.node("doc", null, [
          schema.node(
            "paragraph",
            {
              pPrMark: delMark({ id: 2 }),
              sectionBreakType: "nextPage",
              _sectionProperties: { columns: 2 },
            },
            schema.text("first section"),
          ),
          schema.node(
            "paragraph",
            {
              pPrMark: delMark({ id: 3 }),
              sectionBreakType: "oddPage",
              _sectionProperties: { columns: 3 },
            },
            schema.text("second section"),
          ),
          schema.node(
            "paragraph",
            {
              sectionBreakType: "continuous",
              _sectionProperties: { columns: 4 },
            },
            schema.text("following section"),
          ),
        ]),
      );
      const view = dispatcher(state);

      acceptAIEditRevision(2)(view.state, view.dispatch);
      acceptAIEditRevision(3)(view.state, view.dispatch);

      expect(getTrackedSectionEndpointRemoval(view.state)).toMatchObject({
        sourceParagraphEndpointCount: 3,
        expectedParagraphEndpointCount: 1,
      });
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
