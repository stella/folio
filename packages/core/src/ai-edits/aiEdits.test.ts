import { describe, expect, test } from "bun:test";
import { Schema } from "prosemirror-model";
import { EditorState } from "prosemirror-state";
import type { Transaction } from "prosemirror-state";
import { TableMap } from "prosemirror-tables";

import {
  acceptAIEditRevision,
  acceptAllChanges,
  rejectAIEditRevision,
  rejectAllChanges,
} from "../prosemirror/commands/comments";
import { applyFolioAIEditOperations } from "./apply";
import { getTrackedChangesFromDoc } from "./read";
import { createFolioAIEditSnapshot, createFolioAITextRangeHandle } from "./snapshot";

const schema = new Schema({
  nodes: {
    doc: { content: "block+" },
    paragraph: {
      content: "text*",
      group: "block",
      attrs: {
        listMarker: { default: null },
        pageBreakBefore: { default: null },
        styleId: { default: null },
        // Identity attrs — must NOT be copied when a new block is
        // synthesized from a sibling, otherwise downstream tracking
        // sees duplicate IDs.
        paraId: { default: null },
        textId: { default: null },
        defaultTextFormatting: { default: null },
      },
    },
    text: {},
    table: {
      content: "tableRow+",
      group: "block",
      tableRole: "table",
    },
    tableRow: {
      content: "tableCell*",
      tableRole: "row",
      attrs: {
        trIns: { default: null },
        trDel: { default: null },
        hidden: { default: false },
      },
    },
    tableCell: {
      content: "block+",
      tableRole: "cell",
      attrs: {
        colspan: { default: 1 },
        rowspan: { default: 1 },
        colwidth: { default: null },
        verticalAlign: { default: null },
        tcPrChange: { default: null },
        cellMarker: { default: null },
        _originalFormatting: { default: null },
        _preserveVMergeRestart: { default: null },
        _docxVMergeContinuationCells: { default: null },
      },
    },
  },
  marks: {
    insertion: {
      attrs: {
        revisionId: {},
        author: {},
        date: {},
      },
      toDOM: () => ["ins", 0],
    },
    deletion: {
      attrs: {
        revisionId: {},
        author: {},
        date: {},
      },
      toDOM: () => ["del", 0],
    },
    runPropertyChange: {
      attrs: {
        changes: { default: [] },
      },
      toDOM: () => ["span", 0],
    },
    comment: {
      attrs: {
        commentId: {},
      },
      toDOM: (mark) => ["span", { "data-comment-id": mark.attrs["commentId"] }, 0],
    },
    bold: {
      toDOM: () => ["strong", 0],
    },
    italic: {
      toDOM: () => ["em", 0],
    },
    fontSize: {
      attrs: {
        size: { default: 24 },
      },
      toDOM: (mark) => ["span", { style: `font-size: ${Number(mark.attrs["size"]) / 2}pt` }, 0],
    },
    fontFamily: {
      attrs: {
        ascii: { default: null },
        hAnsi: { default: null },
      },
      toDOM: () => ["span", 0],
    },
    underline: {
      attrs: {
        style: { default: "single" },
      },
      toDOM: () => ["u", 0],
    },
  },
});

type BlockSpec =
  | string
  | {
      text: string;
      listMarker?: string;
      pageBreakBefore?: boolean;
      paraId?: string;
      styleId?: string;
      textId?: string;
    };

const makeState = (blocks: BlockSpec[]) =>
  EditorState.create({
    schema,
    doc: schema.node(
      "doc",
      null,
      blocks.map((block) => {
        const text = typeof block === "string" ? block : block.text;
        const attrs =
          typeof block === "string"
            ? {}
            : {
                listMarker: block.listMarker ?? null,
                pageBreakBefore: block.pageBreakBefore ?? null,
                paraId: block.paraId ?? null,
                styleId: block.styleId ?? null,
                textId: block.textId ?? null,
              };
        return schema.node("paragraph", attrs, text.length === 0 ? [] : [schema.text(text)]);
      }),
    ),
  });

const makeView = (state: EditorState) => {
  const view = {
    state,
    dispatch(transaction: Transaction) {
      view.state = view.state.apply(transaction);
    },
  };
  return view;
};

/**
 * Helpers for the boundary tests below: build a single-paragraph
 * doc whose runs carry pre-existing tracked-change marks. Lets us
 * write declarative input like ["plain", ["del", "shall"], ["ins",
 * "must"], "plain"] instead of hand-rolling schema.text per call.
 */
type PreexistingRun = string | ["ins" | "del", string, number?]; // [kind, text, optional revisionId]

const makeTrackedDoc = (
  runs: PreexistingRun[],
): { state: EditorState; doc: ReturnType<typeof schema.node> } => {
  const insertionType = schema.marks["insertion"];
  const deletionType = schema.marks["deletion"];
  const date = "2026-01-01T00:00:00.000Z";
  const textNodes = runs.flatMap((run) => {
    if (typeof run === "string") {
      return run.length === 0 ? [] : [schema.text(run)];
    }
    const [kind, text, revisionId = 100] = run;
    if (text.length === 0) {
      return [];
    }
    const mark =
      kind === "ins"
        ? insertionType.create({ revisionId, author: "PriorUser", date })
        : deletionType.create({ revisionId, author: "PriorUser", date });
    return [schema.text(text, [mark])];
  });
  const doc = schema.node("doc", null, [schema.node("paragraph", {}, textNodes)]);
  return { state: EditorState.create({ schema, doc }), doc };
};

const collectMarksByText = (state: EditorState): Record<string, string[]> => {
  const marksByText: Record<string, string[]> = {};
  state.doc.descendants((node) => {
    if (!node.isText) {
      return;
    }
    marksByText[node.text ?? ""] = node.marks.map((m) => m.type.name);
  });
  return marksByText;
};

describe("Folio AI edit operations", () => {
  test("reads a vertical merge revision preserved on a collapsed continuation cell", () => {
    const doc = schema.node("doc", null, [
      schema.node("table", null, [
        schema.node("tableRow", null, [
          schema.node(
            "tableCell",
            {
              rowspan: 3,
              _docxVMergeContinuationCells: [
                {
                  type: "tableCell",
                  formatting: { vMerge: "continue" },
                  structuralChange: {
                    type: "tableCellMerge",
                    info: { id: 90, author: "Reviewer", date: "2026-07-16" },
                    verticalMerge: "continue",
                  },
                  content: [{ type: "paragraph", content: [] }],
                },
                {
                  type: "tableCell",
                  formatting: { vMerge: "continue" },
                  structuralChange: {
                    type: "tableCellMerge",
                    info: { id: 90, author: "Reviewer", date: "2026-07-16" },
                    verticalMerge: "continue",
                  },
                  content: [{ type: "paragraph", content: [] }],
                },
              ],
            },
            [schema.node("paragraph", { paraId: "merge-origin" }, [schema.text("Merged content")])],
          ),
        ]),
        schema.node("tableRow"),
        schema.node("tableRow"),
      ]),
    ]);

    expect(getTrackedChangesFromDoc(doc)).toEqual([
      {
        id: 90,
        type: "cellMerged",
        author: "Reviewer",
        date: "2026-07-16",
        text: "Merged content",
        blockId: "merge-origin",
      },
    ]);
  });

  test("rejects invalid text range boundaries", () => {
    expect(
      createFolioAITextRangeHandle({
        blockId: "",
        text: "text",
        startOffset: 0,
        endOffset: 4,
      }),
    ).toBeNull();
    expect(
      createFolioAITextRangeHandle({
        blockId: "block",
        text: "text",
        startOffset: 2,
        endOffset: 2,
      }),
    ).toBeNull();
  });

  test("replaces one exact occurrence through a stable text range", () => {
    const view = makeView(makeState(["repeat repeat"]));
    const snapshot = createFolioAIEditSnapshot(view.state.doc);
    const block = snapshot.blocks.at(0);
    if (!block) {
      throw new Error("expected a block");
    }
    const range = createFolioAITextRangeHandle({
      blockId: block.id,
      text: block.text,
      startOffset: 7,
      endOffset: 13,
    });
    if (!range) {
      throw new Error("expected a range");
    }

    const result = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [{ id: "range-1", type: "replaceRange", range, replace: "done" }],
      mode: "direct",
    });

    expect(result.applied.map(({ id }) => id)).toEqual(["range-1"]);
    expect(view.state.doc.textContent).toBe("repeat done");
  });

  test("comments and formats an exact text range", () => {
    const view = makeView(makeState(["before target after"]));
    const snapshot = createFolioAIEditSnapshot(view.state.doc);
    const block = snapshot.blocks.at(0);
    if (!block) {
      throw new Error("expected a block");
    }
    const range = createFolioAITextRangeHandle({
      blockId: block.id,
      text: block.text,
      startOffset: 7,
      endOffset: 13,
    });
    if (!range) {
      throw new Error("expected a range");
    }

    const result = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [
        { id: "comment", type: "commentOnRange", range, comment: { text: "Review" } },
        { id: "format", type: "formatRange", range, formatting: { bold: true } },
      ],
      mode: "direct",
      createCommentId: () => 42,
    });

    expect(result.applied.map(({ id }) => id).toSorted()).toEqual(["comment", "format"]);
    const target = view.state.doc.nodeAt(8);
    expect(target?.marks.map((mark) => mark.type.name).toSorted()).toEqual(["bold", "comment"]);
  });

  test("tracks range formatting and supports accepting or rejecting it", () => {
    const applyFormatting = () => {
      const view = makeView(makeState(["target"]));
      const snapshot = createFolioAIEditSnapshot(view.state.doc);
      const block = snapshot.blocks.at(0);
      const range = block
        ? createFolioAITextRangeHandle({
            blockId: block.id,
            text: block.text,
            startOffset: 0,
            endOffset: 6,
          })
        : null;
      if (!range) {
        throw new Error("expected a range");
      }
      const result = applyFolioAIEditOperations({
        view,
        snapshot,
        operations: [{ id: "format", type: "formatRange", range, formatting: { italic: true } }],
        author: "Reviewer",
      });
      const revisionId = result.applied.at(0)?.revisionId;
      if (revisionId === undefined) {
        throw new Error("expected a formatting revision");
      }
      return { view, revisionId };
    };

    const accepting = applyFormatting();
    expect(collectMarksByText(accepting.view.state)).toEqual({
      target: ["runPropertyChange", "italic"],
    });
    expect(getTrackedChangesFromDoc(accepting.view.state.doc)).toEqual([
      expect.objectContaining({
        id: accepting.revisionId,
        type: "formatting",
        author: "Reviewer",
        text: "target",
      }),
    ]);
    acceptAIEditRevision(accepting.revisionId)(accepting.view.state, accepting.view.dispatch);
    expect(collectMarksByText(accepting.view.state)).toEqual({ target: ["italic"] });

    const rejecting = applyFormatting();
    rejectAIEditRevision(rejecting.revisionId)(rejecting.view.state, rejecting.view.dispatch);
    expect(collectMarksByText(rejecting.view.state)).toEqual({ target: [] });
  });

  test("does not invent a formatting revision for an unchanged range", () => {
    const view = makeView(makeState(["target"]));
    const italic = schema.marks["italic"]?.create();
    if (!italic) {
      throw new Error("missing italic mark");
    }
    view.dispatch(view.state.tr.addMark(1, 7, italic));
    const snapshot = createFolioAIEditSnapshot(view.state.doc);
    const block = snapshot.blocks.at(0);
    const range = block
      ? createFolioAITextRangeHandle({
          blockId: block.id,
          text: block.text,
          startOffset: 0,
          endOffset: 6,
        })
      : null;
    if (!range) {
      throw new Error("expected a range");
    }
    const result = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [{ id: "format", type: "formatRange", range, formatting: { italic: true } }],
    });
    expect(result.skipped).toEqual([{ id: "format", reason: "noopOperation" }]);
  });

  test("preserves stacked formatting revisions in one batch", () => {
    const applyFormatting = () => {
      const view = makeView(makeState(["target"]));
      const snapshot = createFolioAIEditSnapshot(view.state.doc);
      const block = snapshot.blocks.at(0);
      const range = block
        ? createFolioAITextRangeHandle({
            blockId: block.id,
            text: block.text,
            startOffset: 0,
            endOffset: 6,
          })
        : null;
      if (!range) {
        throw new Error("expected a range");
      }
      const result = applyFolioAIEditOperations({
        view,
        snapshot,
        operations: [
          { id: "bold", type: "formatRange", range, formatting: { bold: true } },
          { id: "italic", type: "formatRange", range, formatting: { italic: true } },
        ],
      });
      expect(result.skipped).toEqual([]);
      expect(result.applied).toHaveLength(2);
      expect(getTrackedChangesFromDoc(view.state.doc)).toHaveLength(2);
      return view;
    };

    const accepting = applyFormatting();
    acceptAllChanges()(accepting.state, accepting.dispatch);
    expect(collectMarksByText(accepting.state)).toEqual({ target: ["bold", "italic"] });

    const rejecting = applyFormatting();
    rejectAllChanges()(rejecting.state, rejecting.dispatch);
    expect(collectMarksByText(rejecting.state)).toEqual({ target: [] });
  });

  test("rejects a range whose selected text hash is stale", () => {
    const view = makeView(makeState(["repeat repeat"]));
    const snapshot = createFolioAIEditSnapshot(view.state.doc);
    const block = snapshot.blocks.at(0);
    if (!block) {
      throw new Error("expected a block");
    }
    const range = createFolioAITextRangeHandle({
      blockId: block.id,
      text: "repeat changed",
      startOffset: 7,
      endOffset: 14,
    });
    if (!range) {
      throw new Error("expected a range");
    }

    const result = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [{ id: "range-1", type: "replaceRange", range, replace: "done" }],
      mode: "direct",
    });

    expect(result.skipped).toEqual([{ id: "range-1", reason: "staleRange" }]);
    expect(view.state.doc.textContent).toBe("repeat repeat");
  });

  test("creates a simple AI-facing block snapshot", () => {
    const state = makeState(["Opening paragraph.", { listMarker: "7.5.1", text: "Payment one." }]);

    const snapshot = createFolioAIEditSnapshot(state.doc);

    expect(snapshot.blocks).toEqual([
      {
        id: "seq-0001",
        kind: "paragraph",
        text: "Opening paragraph.",
      },
      {
        id: "seq-0002",
        kind: "listItem",
        displayLabel: "7.5.1",
        text: "Payment one.",
      },
    ]);
    expect(snapshot.anchors["seq-0001"]?.textHash).toMatch(/^h/u);
  });

  test("an entirely empty document exposes an operation anchor without a visible block", () => {
    const state = makeState([""]);

    const snapshot = createFolioAIEditSnapshot(state.doc);

    expect(snapshot.blocks).toEqual([]);
    expect(snapshot.emptyDocumentAnchorId).toBe("seq-0001");
    expect(snapshot.anchors["seq-0001"]).toMatchObject({
      id: "seq-0001",
      text: "",
      normalizedText: "",
      hashOccurrenceCount: 1,
    });
  });

  test("an empty table cell is not treated as an empty-document operation anchor", () => {
    const doc = schema.node("doc", null, [
      schema.node("table", null, [
        schema.node("tableRow", null, [schema.node("tableCell", null, [schema.node("paragraph")])]),
      ]),
    ]);
    const state = EditorState.create({ schema, doc });

    const snapshot = createFolioAIEditSnapshot(state.doc);

    expect(snapshot.blocks).toEqual([]);
    expect(snapshot.emptyDocumentAnchorId).toBeUndefined();
    expect(snapshot.anchors).toEqual({});
  });

  test("hides text inside a hidden table row from the AI-facing snapshot", () => {
    const doc = schema.node("doc", null, [
      schema.node("table", null, [
        schema.node("tableRow", { hidden: true }, [
          schema.node("tableCell", null, [
            schema.node("paragraph", null, [schema.text("Hidden secret")]),
          ]),
        ]),
        schema.node("tableRow", null, [
          schema.node("tableCell", null, [
            schema.node("paragraph", null, [schema.text("Visible")]),
          ]),
        ]),
      ]),
    ]);
    const state = EditorState.create({ schema, doc });

    const snapshot = createFolioAIEditSnapshot(state.doc);

    expect(snapshot.blocks.map((block) => block.text)).toEqual(["Visible"]);
    expect(
      Object.values(snapshot.anchors).some((anchor) => anchor.text.includes("Hidden secret")),
    ).toBe(false);
  });

  test("snapshot uses sequential fallback ids for duplicate paraIds", () => {
    const state = makeState([
      { text: "First paragraph.", paraId: "AAAA0001" },
      { text: "Second paragraph.", paraId: "AAAA0001" },
      { text: "Third paragraph.", paraId: "BBBB0002" },
    ]);

    const snapshot = createFolioAIEditSnapshot(state.doc);

    expect(snapshot.blocks.map((block) => block.id)).toEqual(["AAAA0001", "seq-0002", "BBBB0002"]);
    expect(snapshot.anchors["AAAA0001"]?.text).toBe("First paragraph.");
    expect(snapshot.anchors["seq-0002"]?.text).toBe("Second paragraph.");
    expect(snapshot.anchors["BBBB0002"]?.text).toBe("Third paragraph.");
  });

  test("captures paragraph style ids in the AI-facing block snapshot", () => {
    const state = makeState([{ styleId: "ClauseHeading1", text: "Payment terms" }]);

    const snapshot = createFolioAIEditSnapshot(state.doc);

    expect(snapshot.blocks[0]).toMatchObject({
      id: "seq-0001",
      styleId: "ClauseHeading1",
      text: "Payment terms",
    });
  });

  test("captures standard heading styles as one-based outline levels", () => {
    const state = makeState([{ styleId: "Heading2", text: "Payment terms" }]);

    expect(createFolioAIEditSnapshot(state.doc).blocks.at(0)).toMatchObject({
      kind: "heading",
      headingLevel: 2,
    });
  });

  test("captures formatted preview runs in the AI-facing block snapshot", () => {
    const boldType = schema.marks["bold"];
    const italicType = schema.marks["italic"];
    const fontSizeType = schema.marks["fontSize"];
    const fontFamilyType = schema.marks["fontFamily"];
    const underlineType = schema.marks["underline"];
    const state = EditorState.create({
      schema,
      doc: schema.node("doc", null, [
        schema.node("paragraph", {}, [
          schema.text("Plain "),
          schema.text("Styled", [
            boldType.create(),
            italicType.create(),
            underlineType.create({ style: "single" }),
            fontSizeType.create({ size: 28 }),
            fontFamilyType.create({ ascii: "Aptos", hAnsi: "Aptos" }),
          ]),
          schema.text(" No underline", [underlineType.create({ style: "none" })]),
        ]),
      ]),
    });

    const snapshot = createFolioAIEditSnapshot(state.doc);

    expect(snapshot.blocks[0]?.text).toBe("Plain Styled No underline");
    expect(snapshot.blocks[0]?.previewRuns).toEqual([
      { text: "Plain " },
      {
        text: "Styled",
        bold: true,
        italic: true,
        underline: true,
        fontFamily: "Aptos",
        fontSizePt: 14,
      },
      { text: " No underline" },
    ]);
  });

  test("does not render explicit underline none as a formatted preview run", () => {
    const state = EditorState.create({
      schema,
      doc: schema.node("doc", null, [
        schema.node("paragraph", { defaultTextFormatting: { underline: { style: "none" } } }, [
          schema.text("Cleared underline"),
        ]),
      ]),
    });

    const snapshot = createFolioAIEditSnapshot(state.doc);

    expect(snapshot.blocks[0]?.previewRuns).toBeUndefined();
  });

  test("applies safe replacements as tracked changes with an attached comment", () => {
    const view = makeView(makeState(["The buyer shall pay."]));
    const snapshot = createFolioAIEditSnapshot(view.state.doc);

    const result = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [
        {
          id: "op-1",
          type: "replaceInBlock",
          blockId: "seq-0001",
          find: "shall",
          replace: "must",
          comment: { text: "Modernised obligation wording." },
        },
      ],
      author: "AI",
      createCommentId: () => 42,
    });

    expect(result.applied).toHaveLength(1);
    expect(result.applied[0]).toMatchObject({ id: "op-1", commentId: 42 });
    expect(typeof result.applied[0]?.revisionId).toBe("number");
    expect(result.skipped).toEqual([]);
    expect(view.state.doc.textContent).toBe("The buyer shallmust pay.");

    const marksByText: Record<string, string[]> = {};
    view.state.doc.descendants((node) => {
      if (!node.isText) {
        return;
      }
      marksByText[node.text ?? ""] = node.marks.map((mark) => mark.type.name);
    });
    expect(marksByText["shall"]).toContain("deletion");
    expect(marksByText["must"]).toContain("insertion");
    expect(marksByText["must"]).toContain("comment");
  });

  test("skips a replacement when the user changed the target block", () => {
    const originalState = makeState(["The buyer shall pay."]);
    const snapshot = createFolioAIEditSnapshot(originalState.doc);
    const view = makeView(makeState(["The buyer must pay."]));

    const result = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [
        {
          id: "op-1",
          type: "replaceInBlock",
          blockId: "seq-0001",
          find: "shall",
          replace: "must",
        },
      ],
    });

    expect(result).toEqual({
      applied: [],
      skipped: [{ id: "op-1", reason: "changedBlock" }],
    });
    expect(view.state.doc.textContent).toBe("The buyer must pay.");
  });

  test("inserts a new block after a list item and inherits its marker attrs", () => {
    const view = makeView(makeState([{ listMarker: "7.5.1", text: "Payment one." }]));
    const snapshot = createFolioAIEditSnapshot(view.state.doc);

    const result = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [
        {
          id: "op-1",
          type: "insertAfterBlock",
          blockId: "seq-0001",
          text: "Payment two.",
          inheritFormatting: true,
        },
      ],
      mode: "direct",
    });

    expect(result.skipped).toEqual([]);
    expect(view.state.doc.childCount).toBe(2);
    expect(view.state.doc.child(1).attrs["listMarker"]).toBe("7.5.1");
    expect(view.state.doc.child(1).textContent).toBe("Payment two.");
  });

  test("does not over-reject when another block shares the same text hash", () => {
    // The freshness gate must check the target block, not the
    // global count of matching hashes — a sibling with the same
    // text getting edited shouldn't skip an op on an unchanged
    // target.
    const originalState = makeState(["Payment.", "Payment.", "Other text."]);
    const snapshot = createFolioAIEditSnapshot(originalState.doc);
    const view = makeView(makeState(["Payment.", "Tweaked.", "Other text."]));

    const result = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [
        {
          id: "op-1",
          type: "replaceInBlock",
          blockId: "seq-0001",
          find: "Payment",
          replace: "Charge",
        },
      ],
      mode: "direct",
    });

    expect(result.applied).toHaveLength(1);
    expect(result.skipped).toEqual([]);
    expect(view.state.doc.firstChild?.textContent).toBe("Charge.");
  });

  test("survives an unrelated insertion before the target block", () => {
    // The resolver must locate the target block by its content
    // signature, not by the snapshot's raw absolute offset — async
    // insertions before the target shift every later position.
    const originalState = makeState(["First paragraph.", "Target paragraph."]);
    const snapshot = createFolioAIEditSnapshot(originalState.doc);
    // Live doc gained a new block at the top after the snapshot.
    const view = makeView(
      makeState([
        "Inserted before the snapshot was taken.",
        "First paragraph.",
        "Target paragraph.",
      ]),
    );

    const result = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [
        {
          id: "op-1",
          type: "replaceInBlock",
          blockId: "seq-0002",
          find: "Target",
          replace: "Renamed",
        },
      ],
      mode: "direct",
    });

    expect(result.applied).toHaveLength(1);
    expect(result.skipped).toEqual([]);
    expect(view.state.doc.child(2).textContent).toBe("Renamed paragraph.");
  });

  test("paraId-anchored op resolves the right block when a same-text duplicate appears before it", () => {
    // Reproduces the chatgpt-codex review concern on #473: with
    // hash+ordinal-only lookup, a duplicate of the snapshot block's
    // text inserted BEFORE the target between snapshot and apply
    // would steal the ordinal and the op would mutate the wrong
    // block. The paraId-direct path keeps the lookup pinned to the
    // originally-referenced paragraph.
    const originalState = makeState([
      { text: "Other paragraph.", paraId: "10000000" },
      { text: "Payment.", paraId: "AAAA0001" },
    ]);
    const snapshot = createFolioAIEditSnapshot(originalState.doc);
    // Live doc has a NEW paragraph with the same text "Payment."
    // inserted BEFORE the original target. Hash+ordinal would now
    // bucket [insertedDup, originalTarget] under the same hash and
    // pick index 0 (the duplicate), mutating the wrong block.
    const view = makeView(
      makeState([
        { text: "Other paragraph.", paraId: "10000000" },
        { text: "Payment.", paraId: "BBBB0002" },
        { text: "Payment.", paraId: "AAAA0001" },
      ]),
    );

    const result = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [
        {
          id: "op-1",
          type: "replaceInBlock",
          blockId: "AAAA0001",
          find: "Payment",
          replace: "Charge",
        },
      ],
      mode: "direct",
    });

    expect(result.applied).toHaveLength(1);
    expect(result.skipped).toEqual([]);
    // The BBBB0002 block (live index 1) must stay untouched.
    expect(view.state.doc.child(1).textContent).toBe("Payment.");
    // The AAAA0001 block (live index 2) is the one that gets edited.
    expect(view.state.doc.child(2).textContent).toBe("Charge.");
  });

  test("paraId-anchored op still skips when the target block changed after snapshot", () => {
    const originalState = makeState([{ text: "Payment.", paraId: "AAAA0001" }]);
    const snapshot = createFolioAIEditSnapshot(originalState.doc);
    const view = makeView(makeState([{ text: "Payment changed.", paraId: "AAAA0001" }]));

    const result = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [
        {
          id: "op-1",
          type: "replaceInBlock",
          blockId: "AAAA0001",
          find: "Payment",
          replace: "Charge",
        },
      ],
      mode: "direct",
    });

    expect(result).toEqual({
      applied: [],
      skipped: [{ id: "op-1", reason: "changedBlock" }],
    });
    expect(view.state.doc.child(0).textContent).toBe("Payment changed.");
  });

  test("serialized precondition detects a stale target with a fresh apply snapshot", () => {
    const originalState = makeState([{ text: "Payment.", paraId: "AAAA0001" }]);
    const originalSnapshot = createFolioAIEditSnapshot(originalState.doc);
    const view = makeView(makeState([{ text: "Payment changed.", paraId: "AAAA0001" }]));
    const currentSnapshot = createFolioAIEditSnapshot(view.state.doc);
    const blockTextHash = originalSnapshot.anchors["AAAA0001"]?.textHash;
    if (blockTextHash === undefined) {
      throw new Error("expected the original block anchor");
    }

    const result = applyFolioAIEditOperations({
      view,
      snapshot: currentSnapshot,
      operations: [
        {
          id: "op-1",
          type: "replaceInBlock",
          blockId: "AAAA0001",
          find: "Payment changed",
          replace: "Charge",
          precondition: { blockTextHash },
        },
      ],
      mode: "direct",
    });

    expect(result).toEqual({
      applied: [],
      skipped: [{ id: "op-1", reason: "preconditionFailed" }],
    });
    expect(view.state.doc.child(0).textContent).toBe("Payment changed.");
  });

  test("paraId-anchored op skips when the live paraId is gone", () => {
    const originalState = makeState([{ text: "Payment.", paraId: "AAAA0001" }]);
    const snapshot = createFolioAIEditSnapshot(originalState.doc);
    const view = makeView(makeState([{ text: "Payment.", paraId: "BBBB0002" }]));

    const result = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [
        {
          id: "op-1",
          type: "replaceInBlock",
          blockId: "AAAA0001",
          find: "Payment",
          replace: "Charge",
        },
      ],
      mode: "direct",
    });

    expect(result).toEqual({
      applied: [],
      skipped: [{ id: "op-1", reason: "missingBlock" }],
    });
    expect(view.state.doc.child(0).textContent).toBe("Payment.");
  });

  test("applies multiple insertAfterBlock ops at the same position in document order", () => {
    // Same-position ops must apply in a deterministic, logical
    // order. Sorting by `from` alone is non-deterministic for
    // ties, and `tr.insert` shifts positions so a later op at the
    // same numeric `from` lands relative to the mutated doc, not
    // the original anchor.
    const view = makeView(makeState(["Anchor block."]));
    const snapshot = createFolioAIEditSnapshot(view.state.doc);

    const result = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [
        {
          id: "op-1",
          type: "insertAfterBlock",
          blockId: "seq-0001",
          text: "First inserted.",
        },
        {
          id: "op-2",
          type: "insertAfterBlock",
          blockId: "seq-0001",
          text: "Second inserted.",
        },
      ],
      mode: "direct",
    });

    expect(result.skipped).toEqual([]);
    expect(view.state.doc.childCount).toBe(3);
    expect(view.state.doc.child(0).textContent).toBe("Anchor block.");
    // Logical order is op-1 first, op-2 second. The bottom-up
    // apply must preserve that.
    expect(view.state.doc.child(1).textContent).toBe("First inserted.");
    expect(view.state.doc.child(2).textContent).toBe("Second inserted.");
  });

  test("inheritFormatting does not copy identity attrs (paraId / textId)", () => {
    // The new block synthesized for an insertAfterBlock with
    // inheritFormatting must keep formatting attrs (listMarker,
    // styleId) from the source but NEVER reuse identity attrs —
    // duplicate paraId/textId values break tracked-change author
    // attribution and any consumer that keys off them.
    const view = makeView(
      makeState([
        {
          text: "Source paragraph.",
          listMarker: "1.",
          paraId: "para-source",
          textId: "text-source",
        },
      ]),
    );
    const snapshot = createFolioAIEditSnapshot(view.state.doc);

    const result = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [
        {
          id: "op-1",
          type: "insertAfterBlock",
          // Source paragraph carries `paraId: "para-source"`, so the
          // snapshot keys it as `para-source` (paraId-anchored).
          blockId: "para-source",
          text: "Inherited follow-up.",
          inheritFormatting: true,
        },
      ],
      mode: "direct",
    });

    expect(result.skipped).toEqual([]);
    const inserted = view.state.doc.child(1);
    expect(inserted.attrs["listMarker"]).toBe("1.");
    expect(inserted.attrs["paraId"]).toBeNull();
    expect(inserted.attrs["textId"]).toBeNull();
  });

  test("replaceBlock with preserveFormatting=false strips block-level attrs", () => {
    // `preserveFormatting=false` lets the model request "drop the
    // list marker, this is just plain text now". The flag is
    // exposed in the operation schema; honour it.
    const view = makeView(makeState([{ listMarker: "1.", text: "List item content." }]));
    const snapshot = createFolioAIEditSnapshot(view.state.doc);

    const result = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [
        {
          id: "op-1",
          type: "replaceBlock",
          blockId: "seq-0001",
          text: "Now just plain text.",
          preserveFormatting: false,
        },
      ],
      mode: "direct",
    });

    expect(result.skipped).toEqual([]);
    expect(view.state.doc.firstChild?.textContent).toBe("Now just plain text.");
    expect(view.state.doc.firstChild?.attrs["listMarker"]).toBeNull();
  });

  test("replaceBlock applies requested styleId in tracked-changes mode", () => {
    const view = makeView(makeState([{ styleId: "BodyText", text: "Intro paragraph." }]));
    const snapshot = createFolioAIEditSnapshot(view.state.doc);

    const result = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [
        {
          id: "op-1",
          type: "replaceBlock",
          blockId: "seq-0001",
          text: "Clause heading.",
          styleId: "ClauseHeading1",
        },
      ],
    });

    expect(result.skipped).toEqual([]);
    expect(result.applied[0]?.revisionIds).toHaveLength(2);
    expect(view.state.doc.firstChild?.attrs["styleId"]).toBe("ClauseHeading1");
  });

  test("replaceBlock marks only diverging tokens, leaves shared runs untouched", () => {
    // The engine should produce a minimal diff: when most words are
    // unchanged, only the changed runs get insertion/deletion marks.
    // A coarse "mark whole block deletion + insert whole block"
    // would tag every shared word and is what we're guarding against.
    const view = makeView(makeState(["The buyer shall pay the seller within thirty days."]));
    const snapshot = createFolioAIEditSnapshot(view.state.doc);

    const result = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [
        {
          id: "op-1",
          type: "replaceBlock",
          blockId: "seq-0001",
          text: "The buyer must pay the seller within sixty days.",
        },
      ],
    });

    expect(result.skipped).toEqual([]);
    const marksByText: Record<string, string[]> = {};
    view.state.doc.descendants((node) => {
      if (!node.isText) {
        return;
      }
      marksByText[node.text ?? ""] = node.marks.map((mark) => mark.type.name);
    });
    // Shared tokens carry no tracked-change marks.
    for (const shared of ["The buyer ", " pay the seller within ", " days."]) {
      expect(marksByText[shared] ?? []).not.toContain("insertion");
      expect(marksByText[shared] ?? []).not.toContain("deletion");
    }
    // Diverging tokens carry exactly the right marks.
    expect(marksByText["shall"]).toContain("deletion");
    expect(marksByText["must"]).toContain("insertion");
    expect(marksByText["thirty"]).toContain("deletion");
    expect(marksByText["sixty"]).toContain("insertion");
  });

  test("snapshot and apply ignore existing tracked-change marks", () => {
    // The AI should see the post-tracked-changes view: existing
    // deletion-marked text is hidden, existing insertion-marked
    // text is included as plain text. Otherwise the model sees
    // "shallmust" smashed together and writes find/replace
    // operations against that confused string.
    //
    // Build a doc whose textContent is "The buyer shallmust pay."
    // (`shall` is a pending deletion, `must` is a pending insertion).
    const insertionMark = schema.marks["insertion"].create({
      revisionId: 1,
      author: "AI",
      date: "2026-01-01T00:00:00.000Z",
    });
    const deletionMark = schema.marks["deletion"].create({
      revisionId: 1,
      author: "AI",
      date: "2026-01-01T00:00:00.000Z",
    });
    const trackedDoc = schema.node("doc", null, [
      schema.node("paragraph", {}, [
        schema.text("The buyer "),
        schema.text("shall", [deletionMark]),
        schema.text("must", [insertionMark]),
        schema.text(" pay."),
      ]),
    ]);
    const view = makeView(EditorState.create({ schema, doc: trackedDoc }));

    const snapshot = createFolioAIEditSnapshot(view.state.doc);
    expect(snapshot.blocks[0]?.text).toBe("The buyer must pay.");

    const result = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [
        {
          id: "op-1",
          type: "replaceInBlock",
          blockId: "seq-0001",
          find: "must",
          replace: "should",
        },
      ],
    });

    expect(result.skipped).toEqual([]);
    // The pending "shall" stays untouched, the pending "must" is
    // now marked deletion in addition to its existing insertion,
    // and "should" sits as the new insertion next to it.
    const marksByText: Record<string, string[]> = {};
    view.state.doc.descendants((node) => {
      if (!node.isText) {
        return;
      }
      marksByText[node.text ?? ""] = node.marks.map((mark) => mark.type.name);
    });
    expect(marksByText["shall"]).toContain("deletion");
    expect(marksByText["must"]).toContain("insertion");
    expect(marksByText["must"]).toContain("deletion");
    expect(marksByText["should"]).toContain("insertion");
  });

  test("AI edit at the start of clean text when block opens with a deletion run", () => {
    // Block live text: "old shall pay." with "old " marked
    // deletion. Clean view: "shall pay." — AI replaces at offset 0.
    // The find lookup must skip the leading deletion run, and the
    // resulting marks must land on the live "shall" position, not
    // the literal first character of the block.
    const { state } = makeTrackedDoc([["del", "old "], "shall pay."]);
    const view = makeView(state);
    const snapshot = createFolioAIEditSnapshot(view.state.doc);

    expect(snapshot.blocks[0]?.text).toBe("shall pay.");

    const result = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [
        {
          id: "op-1",
          type: "replaceInBlock",
          blockId: "seq-0001",
          find: "shall",
          replace: "must",
        },
      ],
    });

    expect(result.skipped).toEqual([]);
    const marks = collectMarksByText(view.state);
    expect(marks["old "]).toContain("deletion");
    expect(marks["shall"]).toContain("deletion");
    expect(marks["must"]).toContain("insertion");
  });

  test("AI edit at the end of clean text when block ends with an insertion run", () => {
    // Block: "Pay " + ins "promptly." → clean view: "Pay promptly."
    // AI replaces "promptly" at the end. Marks must land on the
    // existing insertion run, stacking new del+ins on top.
    const { state } = makeTrackedDoc(["Pay ", ["ins", "promptly."]]);
    const view = makeView(state);
    const snapshot = createFolioAIEditSnapshot(view.state.doc);

    expect(snapshot.blocks[0]?.text).toBe("Pay promptly.");

    const result = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [
        {
          id: "op-1",
          type: "replaceInBlock",
          blockId: "seq-0001",
          find: "promptly",
          replace: "immediately",
        },
      ],
    });

    expect(result.skipped).toEqual([]);
    const marks = collectMarksByText(view.state);
    // "promptly" was already an insertion, now also gets a deletion
    // mark from the new revision; the leading "Pay " stays clean.
    expect(marks["Pay "] ?? []).not.toContain("insertion");
    expect(marks["Pay "] ?? []).not.toContain("deletion");
    expect(marks["promptly"]).toContain("insertion");
    expect(marks["promptly"]).toContain("deletion");
    expect(marks["immediately"]).toContain("insertion");
  });

  test("AI edit spanning across an existing deletion run still resolves", () => {
    // Live runs: "The buyer " + del "shall " + ins "must " + "pay."
    // Clean view: "The buyer must pay." — AI replaces "buyer must"
    // with "seller should". The find spans across the (skipped)
    // "shall" deletion in the live doc; mark boundaries must use
    // the right PM positions and leave "shall" untouched.
    const { state } = makeTrackedDoc(["The buyer ", ["del", "shall "], ["ins", "must "], "pay."]);
    const view = makeView(state);
    const snapshot = createFolioAIEditSnapshot(view.state.doc);

    expect(snapshot.blocks[0]?.text).toBe("The buyer must pay.");

    const result = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [
        {
          id: "op-1",
          type: "replaceInBlock",
          blockId: "seq-0001",
          find: "buyer must",
          replace: "seller should",
        },
      ],
    });

    expect(result.skipped).toEqual([]);
    // Find the node carrying the literal "shall " text (PM may
    // split or re-bucket text nodes after the new addMark calls,
    // so a direct lookup by full text isn't reliable).
    let shallNodeMarks: string[] | null = null;
    view.state.doc.descendants((node) => {
      if (node.isText && node.text !== undefined && node.text.includes("shall")) {
        shallNodeMarks = node.marks.map((m) => m.type.name);
      }
    });
    expect(shallNodeMarks).not.toBeNull();
    // Pre-existing "shall" deletion is left intact (one deletion
    // mark, no insertion mark added on top by the new revision).
    const shallMarks: string[] = shallNodeMarks;
    expect(shallMarks.includes("deletion")).toBe(true);
    expect(shallMarks.includes("insertion")).toBe(false);
  });

  test("a single replace operation uses distinct revisionIds for ins vs del", () => {
    // fromProseDoc's DOCX writer treats any revisionId that appears
    // on BOTH an insertion mark AND a deletion mark in the doc as a
    // Word "move" (w:moveTo / w:moveFrom). For a plain replace
    // operation that's a misclassification — Word would render it
    // as a "moved from / moved to" pair instead of a strike-through
    // + new text. The engine must therefore allocate one revisionId
    // for the deletion side and a different one for the insertion
    // side of a single replace, so they serialise correctly.
    const view = makeView(makeState(["The buyer shall pay."]));
    const snapshot = createFolioAIEditSnapshot(view.state.doc);

    applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [
        {
          id: "op-1",
          type: "replaceInBlock",
          blockId: "seq-0001",
          find: "shall",
          replace: "must",
        },
      ],
    });

    const insertionIds = new Set<number>();
    const deletionIds = new Set<number>();
    view.state.doc.descendants((node) => {
      if (!node.isText) {
        return;
      }
      for (const mark of node.marks) {
        const id = Number(mark.attrs["revisionId"]);
        if (!Number.isFinite(id)) {
          continue;
        }
        if (mark.type.name === "insertion") {
          insertionIds.add(id);
        } else if (mark.type.name === "deletion") {
          deletionIds.add(id);
        }
      }
    });
    const overlap = [...insertionIds].filter((id) => deletionIds.has(id));
    expect(overlap).toEqual([]);
  });

  test("replaceInBlock on a tracked-changes block actually mutates the doc", () => {
    // Regression: when the block carries pending deletion runs,
    // the engine used to slice blockNode.textContent (which
    // includes the deleted chars) using PM positions from the
    // post-tracked-changes view (which doesn't). The mismatch
    // produced a diff with no PM steps, yet `applied[]` still
    // listed the op — a silent accept-failure where the panel
    // said "accepted" but the document didn't change. Lock in
    // that the op is either applied AND the doc mutated, or
    // skipped — never phantom-applied.
    const { state } = makeTrackedDoc(["The buyer ", ["del", "shall "], ["ins", "must "], "pay."]);
    const view = makeView(state);
    const docBefore = view.state.doc.toString();
    const snapshot = createFolioAIEditSnapshot(view.state.doc);

    const result = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [
        {
          id: "op-1",
          type: "replaceInBlock",
          blockId: "seq-0001",
          find: "must",
          replace: "should",
        },
      ],
    });

    expect(result.applied).toHaveLength(1);
    expect(result.skipped).toEqual([]);
    expect(view.state.doc.toString()).not.toBe(docBefore);
  });

  test("filters no-op operations (find equals replace, replaceBlock equals current)", () => {
    // The model occasionally emits replaceInBlock with find ===
    // replace (verified in dev-tools trace). Skip with reason
    // "noopOperation" so the panel never shows X→X cards.
    const view = makeView(makeState(["Prodávající 3."]));
    const snapshot = createFolioAIEditSnapshot(view.state.doc);

    const result = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [
        {
          id: "op-1",
          type: "replaceInBlock",
          blockId: "seq-0001",
          find: "Prodávající 3",
          replace: "Prodávající 3",
        },
        {
          id: "op-2",
          type: "replaceBlock",
          blockId: "seq-0001",
          text: "Prodávající 3.",
        },
      ],
    });

    expect(result.applied).toEqual([]);
    expect(result.skipped).toEqual([
      { id: "op-1", reason: "noopOperation" },
      { id: "op-2", reason: "noopOperation" },
    ]);
  });

  test("two queued ops from one batch survive an in-between accept that mutates structure", () => {
    // Real-world flow: AI generates a single batch with three ops.
    // The user accepts them sequentially. The first accept inserts
    // a paragraph (insertAfterBlock), shifting all subsequent
    // block PM positions; the panel must still be able to resolve
    // ops 2 and 3 against the ORIGINAL snapshot via textHash
    // lookup. A fresh-snapshot-per-accept approach would break op
    // 3 because block ids would have re-numbered.
    const view = makeView(
      makeState(["Section 1 intro.", "Section 2 body.", "Section 3 conclusion."]),
    );
    const originalSnapshot = createFolioAIEditSnapshot(view.state.doc);

    // First op: insert a new paragraph after Section 1. This
    // structurally shifts Section 2 and Section 3 down.
    const r1 = applyFolioAIEditOperations({
      view,
      snapshot: originalSnapshot,
      operations: [
        {
          id: "op-1",
          type: "insertAfterBlock",
          blockId: "seq-0001",
          text: "Inserted aside.",
        },
      ],
      mode: "direct",
    });
    expect(r1.skipped).toEqual([]);
    expect(view.state.doc.childCount).toBe(4);

    // Second op references "seq-0002" in the ORIGINAL snapshot
    // (which was Section 2). After the insertion above, a fresh
    // snapshot would call Section 2 "seq-0003" — but we use the
    // original. The textHash lookup must find Section 2 at its
    // shifted position.
    const r2 = applyFolioAIEditOperations({
      view,
      snapshot: originalSnapshot,
      operations: [
        {
          id: "op-2",
          type: "replaceInBlock",
          blockId: "seq-0002",
          find: "Section 2",
          replace: "Section II",
        },
      ],
      mode: "direct",
    });
    expect(r2.skipped).toEqual([]);
    expect(r2.applied).toHaveLength(1);

    // Third op: also against original snapshot, targets Section 3.
    const r3 = applyFolioAIEditOperations({
      view,
      snapshot: originalSnapshot,
      operations: [
        {
          id: "op-3",
          type: "replaceInBlock",
          blockId: "seq-0003",
          find: "Section 3",
          replace: "Section III",
        },
      ],
      mode: "direct",
    });
    expect(r3.skipped).toEqual([]);
    expect(r3.applied).toHaveLength(1);

    expect(view.state.doc.child(0).textContent).toBe("Section 1 intro.");
    expect(view.state.doc.child(1).textContent).toBe("Inserted aside.");
    expect(view.state.doc.child(2).textContent).toBe("Section II body.");
    expect(view.state.doc.child(3).textContent).toBe("Section III conclusion.");
  });

  test("queued op resolves against the original snapshot after the doc shifts", () => {
    // Locks in why the panel must hand the apply engine the
    // ORIGINAL snapshot the AI saw, not a freshly recomputed one:
    //
    //   1. Fallback block ids are sequential (seq-0001, seq-0002, ...). After
    //      an insertAfterBlock accept, every block below shifts +1.
    //   2. The resolver looks up blocks by `textHash` (content
    //      hash), not by id position. So as long as the target
    //      block's CONTENT hasn't changed since snapshot time,
    //      its hash bucket is unchanged and the lookup succeeds —
    //      even if its absolute PM position moved.
    //   3. A fresh snapshot would re-number the target as a
    //      different id (e.g. seq-0003 → seq-0004), and the queued
    //      op's blockId="seq-0003" would either miss or hit the
    //      wrong block.
    const view = makeView(makeState(["Alpha block.", "Bravo block.", "Charlie block."]));
    const originalSnapshot = createFolioAIEditSnapshot(view.state.doc);

    // Mutate the doc by inserting a new block above Charlie. This
    // shifts Charlie's PM offset but leaves its text (and thus
    // hash) untouched.
    applyFolioAIEditOperations({
      view,
      snapshot: originalSnapshot,
      operations: [
        {
          id: "ins-1",
          type: "insertAfterBlock",
          blockId: "seq-0001",
          text: "Inserted after Alpha.",
        },
      ],
      mode: "direct",
    });
    expect(view.state.doc.childCount).toBe(4);

    // A fresh snapshot would call Charlie "seq-0004" now; the
    // ORIGINAL snapshot still calls it "seq-0003". Apply an op
    // referencing the original id — must succeed against the
    // mutated doc.
    const result = applyFolioAIEditOperations({
      view,
      snapshot: originalSnapshot,
      operations: [
        {
          id: "op-charlie",
          type: "replaceInBlock",
          blockId: "seq-0003",
          find: "Charlie",
          replace: "Delta",
        },
      ],
      mode: "direct",
    });

    expect(result.applied).toHaveLength(1);
    expect(result.skipped).toEqual([]);
    expect(view.state.doc.lastChild?.textContent).toBe("Delta block.");
  });

  test("snapshot omits a block whose entire content is deletion-marked", () => {
    // After the user accepts the deletion the block becomes empty
    // anyway, so the AI has nothing useful to anchor against.
    // Today the snapshot just skips it (zero-length normalized
    // text). Locks in that behaviour.
    const { state } = makeTrackedDoc([["del", "Entire block is gone."]]);
    const snapshot = createFolioAIEditSnapshot(state.doc);
    expect(snapshot.blocks).toHaveLength(0);
  });

  test("word-level diff works for non-Latin scripts split on whitespace", () => {
    // Czech / German / French / Polish all split on whitespace, so
    // the same LCS path applies — only the diverging Czech token
    // should carry tracked-change marks. Locks in that the regex
    // tokeniser handles diacritics without pre/post-processing.
    const view = makeView(makeState(["Kupující musí zaplatit do třiceti dnů."]));
    const snapshot = createFolioAIEditSnapshot(view.state.doc);

    const result = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [
        {
          id: "op-1",
          type: "replaceBlock",
          blockId: "seq-0001",
          text: "Kupující musí zaplatit do šedesáti dnů.",
        },
      ],
    });

    expect(result.skipped).toEqual([]);
    const marksByText: Record<string, string[]> = {};
    view.state.doc.descendants((node) => {
      if (!node.isText) {
        return;
      }
      marksByText[node.text ?? ""] = node.marks.map((mark) => mark.type.name);
    });
    expect(marksByText["třiceti"]).toContain("deletion");
    expect(marksByText["šedesáti"]).toContain("insertion");
    expect(marksByText["Kupující musí zaplatit do "] ?? []).not.toContain("deletion");
  });

  test("insertAfterBlock anchored inside a table cell lands after the table, not in the cell", () => {
    // The AI sometimes anchors an `insertAfterBlock` to a
    // paragraph that lives inside a `tableCell` (e.g. when the
    // last visible block before the desired insertion point is a
    // cell line). Pre-fix, the new block would be inserted inside
    // the cell — visually invisible and structurally wrong. The
    // resolver now escapes outward to the enclosing `table` and
    // places the synthesized sibling adjacent to the table.
    const cellParagraph = schema.node(
      "paragraph",
      { listMarker: null, paraId: "cell-1", textId: null },
      [schema.text("Cell text.")],
    );
    const table = schema.node("table", null, [
      schema.node("tableRow", null, [schema.node("tableCell", null, [cellParagraph])]),
    ]);
    const docNode = schema.node("doc", null, [
      schema.node("paragraph", { listMarker: null, paraId: null, textId: null }, [
        schema.text("Before."),
      ]),
      table,
      schema.node("paragraph", { listMarker: null, paraId: null, textId: null }, [
        schema.text("After."),
      ]),
    ]);
    const state = EditorState.create({ schema, doc: docNode });
    const snapshot = createFolioAIEditSnapshot(state.doc);
    const view = makeView(state);

    const result = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [
        {
          id: "op-1",
          type: "insertAfterBlock",
          blockId: "cell-1",
          text: "Inserted sibling.",
          inheritFormatting: false,
        },
      ],
      mode: "direct",
    });

    expect(result.skipped).toEqual([]);
    expect(result.applied).toHaveLength(1);

    // Doc top-level shape: Before. | table | Inserted sibling. | After.
    expect(view.state.doc.childCount).toBe(4);
    expect(view.state.doc.child(0).textContent).toBe("Before.");
    expect(view.state.doc.child(1).type.name).toBe("table");
    expect(view.state.doc.child(2).type.name).toBe("paragraph");
    expect(view.state.doc.child(2).textContent).toBe("Inserted sibling.");
    expect(view.state.doc.child(3).textContent).toBe("After.");

    // The cell itself must still hold only the original paragraph.
    const liveTable = view.state.doc.child(1);
    const liveRow = liveTable.child(0);
    const liveCell = liveRow.child(0);
    expect(liveCell.childCount).toBe(1);
    expect(liveCell.child(0).textContent).toBe("Cell text.");
  });

  test("insertBeforeBlock anchored inside a table cell lands before the table, not in the cell", () => {
    const cellParagraph = schema.node(
      "paragraph",
      { listMarker: null, paraId: "cell-2", textId: null },
      [schema.text("Cell text.")],
    );
    const table = schema.node("table", null, [
      schema.node("tableRow", null, [schema.node("tableCell", null, [cellParagraph])]),
    ]);
    const docNode = schema.node("doc", null, [
      schema.node("paragraph", { listMarker: null, paraId: null, textId: null }, [
        schema.text("Before."),
      ]),
      table,
    ]);
    const state = EditorState.create({ schema, doc: docNode });
    const snapshot = createFolioAIEditSnapshot(state.doc);
    const view = makeView(state);

    const result = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [
        {
          id: "op-1",
          type: "insertBeforeBlock",
          blockId: "cell-2",
          text: "Pre-table sibling.",
          inheritFormatting: false,
        },
      ],
      mode: "direct",
    });

    expect(result.skipped).toEqual([]);
    expect(view.state.doc.childCount).toBe(3);
    expect(view.state.doc.child(0).textContent).toBe("Before.");
    expect(view.state.doc.child(1).textContent).toBe("Pre-table sibling.");
    expect(view.state.doc.child(2).type.name).toBe("table");
  });

  test("inserted **bold** markdown becomes a real bold run (direct mode)", () => {
    const state = makeState(["Anchor."]);
    const snapshot = createFolioAIEditSnapshot(state.doc);
    const view = makeView(state);

    const result = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [
        {
          id: "op-1",
          type: "insertAfterBlock",
          blockId: snapshot.blocks[0]?.id ?? "",
          text: "**Date:** 2026",
          inheritFormatting: false,
        },
      ],
      mode: "direct",
    });

    expect(result.skipped).toEqual([]);
    const inserted = view.state.doc.child(1);
    expect(inserted.textContent).toBe("Date: 2026");
    const runs: { text: string; marks: string[] }[] = [];
    inserted.descendants((node) => {
      if (node.isText) {
        runs.push({
          text: node.text ?? "",
          marks: node.marks.map((m) => m.type.name),
        });
      }
    });
    expect(runs).toEqual([
      { text: "Date:", marks: ["bold"] },
      { text: " 2026", marks: [] },
    ]);
  });

  test("inserted **bold** markdown carries both insertion and bold marks (tracked changes)", () => {
    const state = makeState(["Anchor."]);
    const snapshot = createFolioAIEditSnapshot(state.doc);
    const view = makeView(state);

    applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [
        {
          id: "op-1",
          type: "insertAfterBlock",
          blockId: snapshot.blocks[0]?.id ?? "",
          text: "**Date:** 2026",
          inheritFormatting: false,
        },
      ],
      mode: "tracked-changes",
    });

    const inserted = view.state.doc.child(1);
    expect(inserted.textContent).toBe("Date: 2026");
    const boldRun = collectMarksByText(view.state)["Date:"];
    expect(boldRun).toContain("insertion");
    expect(boldRun).toContain("bold");
  });

  test("replaceBlock **bold** becomes a real bold run and keeps block attrs (direct mode)", () => {
    const state = makeState([{ text: "Old line.", paraId: "AAAA0001" }]);
    const snapshot = createFolioAIEditSnapshot(state.doc);
    const view = makeView(state);

    const result = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [
        {
          id: "op-1",
          type: "replaceBlock",
          blockId: snapshot.blocks[0]?.id ?? "",
          text: "**Date:** 2026",
        },
      ],
      mode: "direct",
    });

    expect(result.skipped).toEqual([]);
    const block = view.state.doc.child(0);
    expect(block.textContent).toBe("Date: 2026");
    expect(block.attrs["paraId"]).toBe("AAAA0001");
    const runs: { text: string; marks: string[] }[] = [];
    block.descendants((node) => {
      if (node.isText) {
        runs.push({
          text: node.text ?? "",
          marks: node.marks.map((m) => m.type.name),
        });
      }
    });
    expect(runs).toEqual([
      { text: "Date:", marks: ["bold"] },
      { text: " 2026", marks: [] },
    ]);
  });

  test("replaceBlock strips markdown markers in tracked-changes mode (no literal asterisks)", () => {
    const state = makeState([{ text: "Old line.", paraId: "AAAA0001" }]);
    const snapshot = createFolioAIEditSnapshot(state.doc);
    const view = makeView(state);

    applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [
        {
          id: "op-1",
          type: "replaceBlock",
          blockId: snapshot.blocks[0]?.id ?? "",
          text: "**Date:** 2026",
        },
      ],
      mode: "tracked-changes",
    });

    // The word-diff redline interleaves the deleted old text with the new
    // run, so the new text is present but not contiguous. The point here is
    // that the markdown markers never reach the document as literal text.
    const text = view.state.doc.textContent;
    expect(text).toContain("Date:");
    expect(text).toContain("2026");
    expect(text).not.toContain("*");
  });

  test("replaceInBlock **bold** becomes a real bold run mid-block (direct mode)", () => {
    const state = makeState(["Effective Date: TBD."]);
    const snapshot = createFolioAIEditSnapshot(state.doc);
    const view = makeView(state);

    const result = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [
        {
          id: "op-1",
          type: "replaceInBlock",
          blockId: snapshot.blocks[0]?.id ?? "",
          find: "Date:",
          replace: "**Date:**",
        },
      ],
      mode: "direct",
    });

    expect(result.skipped).toEqual([]);
    const block = view.state.doc.child(0);
    expect(block.textContent).toBe("Effective Date: TBD.");
    expect(block.textContent).not.toContain("*");
    const runs: { text: string; marks: string[] }[] = [];
    block.descendants((node) => {
      if (node.isText) {
        runs.push({
          text: node.text ?? "",
          marks: node.marks.map((m) => m.type.name),
        });
      }
    });
    expect(runs).toEqual([
      { text: "Effective ", marks: [] },
      { text: "Date:", marks: ["bold"] },
      { text: " TBD.", marks: [] },
    ]);
  });

  test("page-break-only inserts are skipped in tracked-changes mode", () => {
    const view = makeView(makeState(["Anchor block."]));
    const snapshot = createFolioAIEditSnapshot(view.state.doc);

    const result = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [
        {
          id: "op-1",
          type: "insertAfterBlock",
          blockId: "seq-0001",
          text: "",
          pageBreakBefore: true,
        },
      ],
    });

    expect(result.applied).toEqual([]);
    expect(result.skipped).toEqual([{ id: "op-1", reason: "unsupportedMode" }]);
    expect(view.state.doc.childCount).toBe(1);
  });

  test("signature-table inserts are skipped in tracked-changes mode", () => {
    const view = makeView(makeState(["Anchor block."]));
    const snapshot = createFolioAIEditSnapshot(view.state.doc);

    const result = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [
        {
          id: "op-1",
          type: "insertSignatureTable",
          blockId: "seq-0001",
          parties: [{ name: "Buyer" }, { name: "Seller" }],
        },
      ],
    });

    expect(result.applied).toEqual([]);
    expect(result.skipped).toEqual([{ id: "op-1", reason: "unsupportedMode" }]);
    expect(view.state.doc.childCount).toBe(1);
  });

  test("signature-table inserts anchored inside a table cell land after the table", () => {
    const cellParagraph = schema.node(
      "paragraph",
      { listMarker: null, paraId: "cell-3", textId: null },
      [schema.text("Cell text.")],
    );
    const table = schema.node("table", null, [
      schema.node("tableRow", null, [schema.node("tableCell", null, [cellParagraph])]),
    ]);
    const docNode = schema.node("doc", null, [
      schema.node("paragraph", { listMarker: null, paraId: null, textId: null }, [
        schema.text("Before."),
      ]),
      table,
      schema.node("paragraph", { listMarker: null, paraId: null, textId: null }, [
        schema.text("After."),
      ]),
    ]);
    const state = EditorState.create({ schema, doc: docNode });
    const snapshot = createFolioAIEditSnapshot(state.doc);
    const view = makeView(state);

    const result = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [
        {
          id: "op-1",
          type: "insertSignatureTable",
          blockId: "cell-3",
          parties: [{ name: "Buyer" }, { name: "Seller" }],
        },
      ],
      mode: "direct",
    });

    expect(result.skipped).toEqual([]);
    expect(result.applied).toHaveLength(1);
    expect(view.state.doc.childCount).toBe(4);
    expect(view.state.doc.child(0).textContent).toBe("Before.");
    expect(view.state.doc.child(1).type.name).toBe("table");
    expect(view.state.doc.child(2).type.name).toBe("table");
    expect(view.state.doc.child(2).textContent).toContain("Buyer");
    expect(view.state.doc.child(3).textContent).toBe("After.");
  });

  test("inserts ordered rows while preserving cells that span the insertion boundary", () => {
    const spanningCell = schema.node("tableCell", { colspan: 2, rowspan: 2 }, [
      schema.node("paragraph", { paraId: "cell-a" }, [schema.text("A")]),
    ]);
    const table = schema.node("table", null, [
      schema.node("tableRow", null, [
        spanningCell,
        schema.node("tableCell", null, [
          schema.node("paragraph", { paraId: "cell-b" }, [schema.text("B")]),
        ]),
      ]),
      schema.node("tableRow", null, [
        schema.node("tableCell", null, [
          schema.node("paragraph", { paraId: "cell-c" }, [schema.text("C")]),
        ]),
      ]),
    ]);
    const state = EditorState.create({ schema, doc: schema.node("doc", null, [table]) });
    const snapshot = createFolioAIEditSnapshot(state.doc);
    const view = makeView(state);

    const result = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [
        {
          id: "insert-first-row",
          type: "insertTableRow",
          blockId: "cell-c",
          position: "before",
          cellTexts: ["First"],
        },
        {
          id: "insert-second-row",
          type: "insertTableRow",
          blockId: "cell-c",
          position: "before",
          cellTexts: ["Second"],
        },
      ],
      mode: "direct",
    });

    expect(result.skipped).toEqual([]);
    expect(result.applied.map(({ id }) => id).toSorted()).toEqual([
      "insert-first-row",
      "insert-second-row",
    ]);
    const updatedTable = view.state.doc.child(0);
    expect(updatedTable.childCount).toBe(4);
    expect(updatedTable.child(0).child(0).attrs["rowspan"]).toBe(4);
    expect(updatedTable.child(1).childCount).toBe(1);
    expect(updatedTable.child(1).textContent).toBe("First");
    expect(updatedTable.child(2).textContent).toBe("Second");
    expect(updatedTable.child(3).textContent).toBe("C");
  });

  test("targets the nearest row when the anchor is inside a nested table", () => {
    const innerTable = schema.node("table", null, [
      schema.node("tableRow", null, [
        schema.node("tableCell", null, [
          schema.node("paragraph", { paraId: "inner-cell" }, [schema.text("Inner")]),
        ]),
      ]),
    ]);
    const outerTable = schema.node("table", null, [
      schema.node("tableRow", null, [
        schema.node("tableCell", null, [
          schema.node("paragraph", null, [schema.text("Outer")]),
          innerTable,
        ]),
      ]),
    ]);
    const state = EditorState.create({ schema, doc: schema.node("doc", null, [outerTable]) });
    const snapshot = createFolioAIEditSnapshot(state.doc);
    const view = makeView(state);

    const result = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [
        {
          id: "insert-inner-row",
          type: "insertTableRow",
          blockId: "inner-cell",
          cellTexts: ["New inner"],
        },
      ],
      mode: "direct",
    });

    expect(result.skipped).toEqual([]);
    const updatedOuterTable = view.state.doc.child(0);
    expect(updatedOuterTable.childCount).toBe(1);
    const updatedInnerTable = updatedOuterTable.child(0).child(0).child(1);
    expect(updatedInnerTable.childCount).toBe(2);
    expect(updatedInnerTable.child(1).textContent).toBe("New inner");
  });

  test("deletes rows while preserving cells that span the deletion boundary", () => {
    const makeMergedTableState = () => {
      const spanningCell = schema.node("tableCell", { colspan: 2, rowspan: 2 }, [
        schema.node("paragraph", { paraId: "delete-a" }, [schema.text("A")]),
      ]);
      const table = schema.node("table", null, [
        schema.node("tableRow", null, [
          spanningCell,
          schema.node("tableCell", null, [
            schema.node("paragraph", { paraId: "delete-b" }, [schema.text("B")]),
          ]),
        ]),
        schema.node("tableRow", null, [
          schema.node("tableCell", null, [
            schema.node("paragraph", { paraId: "delete-c" }, [schema.text("C")]),
          ]),
        ]),
      ]);
      return EditorState.create({ schema, doc: schema.node("doc", null, [table]) });
    };

    const originState = makeMergedTableState();
    const originView = makeView(originState);
    const originResult = applyFolioAIEditOperations({
      view: originView,
      snapshot: createFolioAIEditSnapshot(originState.doc),
      operations: [{ id: "delete-origin", type: "deleteTableRow", blockId: "delete-a" }],
      mode: "direct",
    });

    expect(originResult.skipped).toEqual([]);
    const originTable = originView.state.doc.child(0);
    expect(originTable.childCount).toBe(1);
    expect(originTable.child(0).childCount).toBe(2);
    expect(originTable.child(0).child(0).attrs["colspan"]).toBe(2);
    expect(originTable.child(0).child(0).attrs["rowspan"]).toBe(1);
    expect(originTable.textContent).toBe("AC");

    const continuationState = makeMergedTableState();
    const continuationView = makeView(continuationState);
    const continuationResult = applyFolioAIEditOperations({
      view: continuationView,
      snapshot: createFolioAIEditSnapshot(continuationState.doc),
      operations: [{ id: "delete-continuation", type: "deleteTableRow", blockId: "delete-c" }],
      mode: "direct",
    });

    expect(continuationResult.skipped).toEqual([]);
    const continuationTable = continuationView.state.doc.child(0);
    expect(continuationTable.childCount).toBe(1);
    expect(continuationTable.child(0).childCount).toBe(2);
    expect(continuationTable.child(0).child(0).attrs["rowspan"]).toBe(1);
    expect(continuationTable.textContent).toBe("AB");
  });

  test("deletes only the nearest table when the anchor is nested", () => {
    const innerTable = schema.node("table", null, [
      schema.node("tableRow", null, [
        schema.node("tableCell", null, [
          schema.node("paragraph", { paraId: "delete-inner" }, [schema.text("Inner")]),
        ]),
      ]),
    ]);
    const outerTable = schema.node("table", null, [
      schema.node("tableRow", null, [
        schema.node("tableCell", null, [
          schema.node("paragraph", null, [schema.text("Before")]),
          innerTable,
          schema.node("paragraph", null, [schema.text("After")]),
        ]),
      ]),
    ]);
    const state = EditorState.create({ schema, doc: schema.node("doc", null, [outerTable]) });
    const view = makeView(state);

    const result = applyFolioAIEditOperations({
      view,
      snapshot: createFolioAIEditSnapshot(state.doc),
      operations: [{ id: "delete-inner-row", type: "deleteTableRow", blockId: "delete-inner" }],
      mode: "direct",
    });

    expect(result.skipped).toEqual([]);
    expect(view.state.doc.childCount).toBe(1);
    const updatedOuterCell = view.state.doc.child(0).child(0).child(0);
    expect(updatedOuterCell.childCount).toBe(2);
    expect(updatedOuterCell.textContent).toBe("BeforeAfter");
  });

  test("keeps a valid parent when deleting the only row from the only table", () => {
    const table = schema.node("table", null, [
      schema.node("tableRow", null, [
        schema.node("tableCell", null, [
          schema.node("paragraph", { paraId: "delete-only" }, [schema.text("Only")]),
        ]),
      ]),
    ]);
    const state = EditorState.create({ schema, doc: schema.node("doc", null, [table]) });
    const view = makeView(state);

    const result = applyFolioAIEditOperations({
      view,
      snapshot: createFolioAIEditSnapshot(state.doc),
      operations: [{ id: "delete-only-row", type: "deleteTableRow", blockId: "delete-only" }],
      mode: "direct",
    });

    expect(result.applied).toEqual([{ id: "delete-only-row" }]);
    expect(result.skipped).toEqual([]);
    expect(view.state.doc.childCount).toBe(1);
    expect(view.state.doc.child(0).type.name).toBe("paragraph");
    expect(view.state.doc.textContent).toBe("");
  });

  test("does not delete an extra row when a batch anchors the same row twice", () => {
    const table = schema.node("table", null, [
      schema.node("tableRow", null, [
        schema.node("tableCell", null, [
          schema.node("paragraph", { paraId: "same-row-a" }, [schema.text("A")]),
        ]),
        schema.node("tableCell", null, [
          schema.node("paragraph", { paraId: "same-row-b" }, [schema.text("B")]),
        ]),
      ]),
      schema.node("tableRow", null, [
        schema.node("tableCell", null, [
          schema.node("paragraph", { paraId: "other-row" }, [schema.text("C")]),
        ]),
      ]),
    ]);
    const state = EditorState.create({ schema, doc: schema.node("doc", null, [table]) });
    const view = makeView(state);

    const result = applyFolioAIEditOperations({
      view,
      snapshot: createFolioAIEditSnapshot(state.doc),
      operations: [
        { id: "delete-row-a", type: "deleteTableRow", blockId: "same-row-a" },
        { id: "delete-row-b", type: "deleteTableRow", blockId: "same-row-b" },
      ],
      mode: "direct",
    });

    expect(result.applied).toEqual([{ id: "delete-row-a" }]);
    expect(result.skipped).toEqual([{ id: "delete-row-b", reason: "noopOperation" }]);
    expect(view.state.doc.child(0).childCount).toBe(1);
    expect(view.state.doc.textContent).toBe("C");
  });

  test("recomputes table topology between row deletions in one batch", () => {
    const rows = ["A", "B", "C"].map((text) =>
      schema.node("tableRow", null, [
        schema.node("tableCell", null, [
          schema.node("paragraph", { paraId: `batch-${text}` }, [schema.text(text)]),
        ]),
      ]),
    );
    const state = EditorState.create({
      schema,
      doc: schema.node("doc", null, [schema.node("table", null, rows)]),
    });
    const view = makeView(state);

    const result = applyFolioAIEditOperations({
      view,
      snapshot: createFolioAIEditSnapshot(state.doc),
      operations: [
        { id: "delete-middle", type: "deleteTableRow", blockId: "batch-B" },
        { id: "delete-last", type: "deleteTableRow", blockId: "batch-C" },
      ],
      mode: "direct",
    });

    expect(result.skipped).toEqual([]);
    expect(result.applied.map(({ id }) => id).toSorted()).toEqual(["delete-last", "delete-middle"]);
    expect(view.state.doc.child(0).childCount).toBe(1);
    expect(view.state.doc.textContent).toBe("A");
  });

  test("inserts a column while preserving cells that span its boundary", () => {
    const table = schema.node("table", null, [
      schema.node("tableRow", null, [
        schema.node("tableCell", { colspan: 2 }, [
          schema.node("paragraph", { paraId: "column-span" }, [schema.text("A")]),
        ]),
      ]),
      schema.node("tableRow", null, [
        schema.node("tableCell", null, [
          schema.node("paragraph", { paraId: "column-left" }, [schema.text("B")]),
        ]),
        schema.node("tableCell", null, [
          schema.node("paragraph", { paraId: "column-right" }, [schema.text("C")]),
        ]),
      ]),
    ]);
    const state = EditorState.create({ schema, doc: schema.node("doc", null, [table]) });
    const view = makeView(state);

    const result = applyFolioAIEditOperations({
      view,
      snapshot: createFolioAIEditSnapshot(state.doc),
      operations: [
        {
          id: "insert-column",
          type: "insertTableColumn",
          blockId: "column-left",
          cellTexts: ["New"],
        },
      ],
      mode: "direct",
    });

    expect(result.applied).toEqual([{ id: "insert-column" }]);
    expect(result.skipped).toEqual([]);
    const updatedTable = view.state.doc.child(0);
    expect(TableMap.get(updatedTable).width).toBe(3);
    expect(updatedTable.child(0).childCount).toBe(1);
    expect(updatedTable.child(0).child(0).attrs["colspan"]).toBe(3);
    expect(updatedTable.child(1).childCount).toBe(3);
    expect(updatedTable.child(1).textContent).toBe("BNewC");
  });

  test("rejects excess column cell text without mutating the table", () => {
    const table = schema.node("table", null, [
      schema.node("tableRow", null, [
        schema.node("tableCell", { colspan: 2 }, [
          schema.node("paragraph", { paraId: "column-span-top" }, [schema.text("A")]),
        ]),
      ]),
      schema.node("tableRow", null, [
        schema.node("tableCell", null, [
          schema.node("paragraph", { paraId: "column-anchor" }, [schema.text("B")]),
        ]),
        schema.node("tableCell", null, [schema.node("paragraph", null, [schema.text("C")])]),
      ]),
    ]);
    const state = EditorState.create({ schema, doc: schema.node("doc", null, [table]) });
    const view = makeView(state);

    const result = applyFolioAIEditOperations({
      view,
      snapshot: createFolioAIEditSnapshot(state.doc),
      operations: [
        {
          id: "insert-column",
          type: "insertTableColumn",
          blockId: "column-anchor",
          cellTexts: ["One", "Two"],
        },
      ],
      mode: "direct",
    });

    expect(result.applied).toEqual([]);
    expect(result.skipped).toEqual([{ id: "insert-column", reason: "unsupportedBlock" }]);
    expect(view.state.doc).toEqual(state.doc);
  });

  test("targets the nearest column when the anchor is inside a nested table", () => {
    const innerTable = schema.node("table", null, [
      schema.node("tableRow", null, [
        schema.node("tableCell", null, [
          schema.node("paragraph", { paraId: "inner-column" }, [schema.text("Inner")]),
        ]),
      ]),
    ]);
    const outerTable = schema.node("table", null, [
      schema.node("tableRow", null, [
        schema.node("tableCell", null, [
          schema.node("paragraph", null, [schema.text("Outer")]),
          innerTable,
        ]),
      ]),
    ]);
    const state = EditorState.create({ schema, doc: schema.node("doc", null, [outerTable]) });
    const view = makeView(state);

    const result = applyFolioAIEditOperations({
      view,
      snapshot: createFolioAIEditSnapshot(state.doc),
      operations: [
        {
          id: "insert-inner-column",
          type: "insertTableColumn",
          blockId: "inner-column",
          cellTexts: ["New inner"],
        },
      ],
      mode: "direct",
    });

    expect(result.skipped).toEqual([]);
    const updatedOuterTable = view.state.doc.child(0);
    expect(updatedOuterTable.child(0).childCount).toBe(1);
    const updatedInnerTable = updatedOuterTable.child(0).child(0).child(1);
    expect(updatedInnerTable.child(0).childCount).toBe(2);
    expect(updatedInnerTable.textContent).toBe("InnerNew inner");
  });

  test("skips column insertion when a cell has no table ancestors", () => {
    const shallowCell = schema.nodes.tableCell.create(null, [
      schema.node("paragraph", { paraId: "shallow-column" }, [schema.text("Cell")]),
    ]);
    const malformedDoc = schema.nodes.doc.create(null, [shallowCell]);
    const state = EditorState.create({ schema, doc: malformedDoc });
    const view = makeView(state);

    const result = applyFolioAIEditOperations({
      view,
      snapshot: createFolioAIEditSnapshot(state.doc),
      operations: [
        {
          id: "insert-column",
          type: "insertTableColumn",
          blockId: "shallow-column",
        },
      ],
      mode: "direct",
    });

    expect(result).toEqual({
      applied: [],
      skipped: [{ id: "insert-column", reason: "unsupportedBlock" }],
    });
    expect(view.state.doc).toEqual(malformedDoc);
  });

  test("orders same-boundary and distinct column inserts against the original table", () => {
    const makeTableState = () => {
      const rows = [
        ["A", "B"],
        ["C", "D"],
      ].map((texts) =>
        schema.node(
          "tableRow",
          null,
          texts.map((text) =>
            schema.node("tableCell", null, [
              schema.node("paragraph", { paraId: `column-${text}` }, [schema.text(text)]),
            ]),
          ),
        ),
      );
      return EditorState.create({
        schema,
        doc: schema.node("doc", null, [schema.node("table", null, rows)]),
      });
    };

    const sameBoundaryState = makeTableState();
    const sameBoundaryView = makeView(sameBoundaryState);
    const sameBoundaryResult = applyFolioAIEditOperations({
      view: sameBoundaryView,
      snapshot: createFolioAIEditSnapshot(sameBoundaryState.doc),
      operations: [
        {
          id: "first-column",
          type: "insertTableColumn",
          blockId: "column-A",
          cellTexts: ["First top", "First bottom"],
        },
        {
          id: "second-column",
          type: "insertTableColumn",
          blockId: "column-C",
          cellTexts: ["Second top", "Second bottom"],
        },
      ],
      mode: "direct",
    });

    expect(sameBoundaryResult.skipped).toEqual([]);
    expect(sameBoundaryView.state.doc.child(0).child(0).textContent).toBe("AFirst topSecond topB");
    expect(sameBoundaryView.state.doc.child(0).child(1).textContent).toBe(
      "CFirst bottomSecond bottomD",
    );

    const distinctState = makeTableState();
    const distinctView = makeView(distinctState);
    const distinctResult = applyFolioAIEditOperations({
      view: distinctView,
      snapshot: createFolioAIEditSnapshot(distinctState.doc),
      operations: [
        {
          id: "left-column",
          type: "insertTableColumn",
          blockId: "column-C",
          cellTexts: ["Left top", "Left bottom"],
        },
        {
          id: "right-column",
          type: "insertTableColumn",
          blockId: "column-B",
          cellTexts: ["Right top", "Right bottom"],
        },
      ],
      mode: "direct",
    });

    expect(distinctResult.skipped).toEqual([]);
    expect(distinctView.state.doc.child(0).child(0).textContent).toBe("ALeft topBRight top");
    expect(distinctView.state.doc.child(0).child(1).textContent).toBe("CLeft bottomDRight bottom");
  });

  test("maps column insertion through earlier operations in a mixed batch", () => {
    const rows = [
      ["A", "B"],
      ["C", "D"],
    ].map((texts) =>
      schema.node(
        "tableRow",
        null,
        texts.map((text) =>
          schema.node("tableCell", null, [
            schema.node("paragraph", { paraId: `mixed-${text}` }, [schema.text(text)]),
          ]),
        ),
      ),
    );
    const state = EditorState.create({
      schema,
      doc: schema.node("doc", null, [schema.node("table", null, rows)]),
    });
    const view = makeView(state);

    const result = applyFolioAIEditOperations({
      view,
      snapshot: createFolioAIEditSnapshot(state.doc),
      operations: [
        {
          id: "insert-column",
          type: "insertTableColumn",
          blockId: "mixed-C",
          cellTexts: ["New top", "New bottom"],
        },
        {
          id: "insert-row",
          type: "insertTableRow",
          blockId: "mixed-C",
          cellTexts: ["E", "F"],
        },
        {
          id: "insert-before-table",
          type: "insertBeforeBlock",
          blockId: "mixed-A",
          text: "Before table",
        },
      ],
      mode: "direct",
    });

    expect(result.skipped).toEqual([]);
    expect(view.state.doc.child(0).textContent).toBe("Before table");
    const updatedTable = view.state.doc.child(1);
    expect(updatedTable.childCount).toBe(3);
    expect(updatedTable.child(0).textContent).toBe("ANew topB");
    expect(updatedTable.child(1).textContent).toBe("CNew bottomD");
    expect(updatedTable.child(2).childCount).toBe(3);
    expect(updatedTable.child(2).textContent).toBe("EF");
  });

  test("deletes a column while preserving cells that span it", () => {
    const table = schema.node("table", null, [
      schema.node("tableRow", null, [
        schema.node("tableCell", { colspan: 2, colwidth: [100, 200] }, [
          schema.node("paragraph", { paraId: "delete-column-span" }, [schema.text("A")]),
        ]),
        schema.node("tableCell", null, [schema.node("paragraph", null, [schema.text("B")])]),
      ]),
      schema.node("tableRow", null, [
        schema.node("tableCell", null, [
          schema.node("paragraph", { paraId: "delete-column-left" }, [schema.text("C")]),
        ]),
        schema.node("tableCell", null, [
          schema.node("paragraph", { paraId: "delete-column-middle" }, [schema.text("D")]),
        ]),
        schema.node("tableCell", null, [schema.node("paragraph", null, [schema.text("E")])]),
      ]),
    ]);
    const state = EditorState.create({ schema, doc: schema.node("doc", null, [table]) });
    const view = makeView(state);

    const result = applyFolioAIEditOperations({
      view,
      snapshot: createFolioAIEditSnapshot(state.doc),
      operations: [
        {
          id: "delete-column",
          type: "deleteTableColumn",
          blockId: "delete-column-middle",
        },
      ],
      mode: "direct",
    });

    expect(result).toEqual({ applied: [{ id: "delete-column" }], skipped: [] });
    const updatedTable = view.state.doc.child(0);
    expect(TableMap.get(updatedTable).width).toBe(2);
    expect(updatedTable.child(0).child(0).attrs["colspan"]).toBe(1);
    expect(updatedTable.child(0).child(0).attrs["colwidth"]).toEqual([100]);
    expect(updatedTable.child(0).textContent).toBe("AB");
    expect(updatedTable.child(1).textContent).toBe("CE");
  });

  test("targets the nearest column for deletion inside a nested table", () => {
    const innerTable = schema.node("table", null, [
      schema.node("tableRow", null, [
        schema.node("tableCell", null, [
          schema.node("paragraph", { paraId: "delete-inner-column" }, [schema.text("Inner A")]),
        ]),
        schema.node("tableCell", null, [schema.node("paragraph", null, [schema.text("Inner B")])]),
      ]),
    ]);
    const outerTable = schema.node("table", null, [
      schema.node("tableRow", null, [
        schema.node("tableCell", null, [
          schema.node("paragraph", null, [schema.text("Outer")]),
          innerTable,
        ]),
      ]),
    ]);
    const state = EditorState.create({ schema, doc: schema.node("doc", null, [outerTable]) });
    const view = makeView(state);

    const result = applyFolioAIEditOperations({
      view,
      snapshot: createFolioAIEditSnapshot(state.doc),
      operations: [
        {
          id: "delete-inner-column",
          type: "deleteTableColumn",
          blockId: "delete-inner-column",
        },
      ],
      mode: "direct",
    });

    expect(result.skipped).toEqual([]);
    const updatedOuterTable = view.state.doc.child(0);
    expect(updatedOuterTable.child(0).childCount).toBe(1);
    const updatedInnerTable = updatedOuterTable.child(0).child(0).child(1);
    expect(TableMap.get(updatedInnerTable).width).toBe(1);
    expect(updatedInnerTable.textContent).toBe("Inner B");
  });

  test("keeps a valid parent when deleting the only table column", () => {
    const table = schema.node("table", null, [
      schema.node("tableRow", null, [
        schema.node("tableCell", null, [
          schema.node("paragraph", { paraId: "delete-only-column" }, [schema.text("Cell")]),
        ]),
      ]),
    ]);
    const state = EditorState.create({ schema, doc: schema.node("doc", null, [table]) });
    const view = makeView(state);

    const result = applyFolioAIEditOperations({
      view,
      snapshot: createFolioAIEditSnapshot(state.doc),
      operations: [
        {
          id: "delete-column",
          type: "deleteTableColumn",
          blockId: "delete-only-column",
        },
      ],
      mode: "direct",
    });

    expect(result).toEqual({ applied: [{ id: "delete-column" }], skipped: [] });
    expect(view.state.doc.childCount).toBe(1);
    expect(view.state.doc.child(0).type.name).toBe("paragraph");
    expect(view.state.doc.textContent).toBe("");
  });

  test("does not delete an extra column when a batch targets the same column twice", () => {
    const rows = [
      ["A", "B"],
      ["C", "D"],
    ].map((texts) =>
      schema.node(
        "tableRow",
        null,
        texts.map((text) =>
          schema.node("tableCell", null, [
            schema.node("paragraph", { paraId: `duplicate-column-${text}` }, [schema.text(text)]),
          ]),
        ),
      ),
    );
    const state = EditorState.create({
      schema,
      doc: schema.node("doc", null, [schema.node("table", null, rows)]),
    });
    const view = makeView(state);

    const result = applyFolioAIEditOperations({
      view,
      snapshot: createFolioAIEditSnapshot(state.doc),
      operations: [
        { id: "delete-top", type: "deleteTableColumn", blockId: "duplicate-column-A" },
        { id: "delete-bottom", type: "deleteTableColumn", blockId: "duplicate-column-C" },
      ],
      mode: "direct",
    });

    expect(result.applied).toEqual([{ id: "delete-top" }]);
    expect(result.skipped).toEqual([{ id: "delete-bottom", reason: "noopOperation" }]);
    expect(TableMap.get(view.state.doc.child(0)).width).toBe(1);
    expect(view.state.doc.child(0).textContent).toBe("BD");
  });

  test("recomputes table topology between column deletions in one batch", () => {
    const rows = [
      ["A", "B", "C"],
      ["D", "E", "F"],
    ].map((texts) =>
      schema.node(
        "tableRow",
        null,
        texts.map((text) =>
          schema.node("tableCell", null, [
            schema.node("paragraph", { paraId: `batch-column-${text}` }, [schema.text(text)]),
          ]),
        ),
      ),
    );
    const state = EditorState.create({
      schema,
      doc: schema.node("doc", null, [schema.node("table", null, rows)]),
    });
    const view = makeView(state);

    const result = applyFolioAIEditOperations({
      view,
      snapshot: createFolioAIEditSnapshot(state.doc),
      operations: [
        { id: "delete-middle", type: "deleteTableColumn", blockId: "batch-column-B" },
        { id: "delete-right", type: "deleteTableColumn", blockId: "batch-column-C" },
      ],
      mode: "direct",
    });

    expect(result.skipped).toEqual([]);
    expect(result.applied.map(({ id }) => id).toSorted()).toEqual([
      "delete-middle",
      "delete-right",
    ]);
    expect(TableMap.get(view.state.doc.child(0)).width).toBe(1);
    expect(view.state.doc.child(0).textContent).toBe("AD");
  });

  test("keeps insertion and deletion targets stable at the same column boundary", () => {
    const rows = [
      ["A", "B"],
      ["C", "D"],
    ].map((texts) =>
      schema.node(
        "tableRow",
        null,
        texts.map((text) =>
          schema.node("tableCell", null, [
            schema.node("paragraph", { paraId: `mixed-column-${text}` }, [schema.text(text)]),
          ]),
        ),
      ),
    );
    const state = EditorState.create({
      schema,
      doc: schema.node("doc", null, [schema.node("table", null, rows)]),
    });
    const view = makeView(state);

    const result = applyFolioAIEditOperations({
      view,
      snapshot: createFolioAIEditSnapshot(state.doc),
      operations: [
        { id: "delete-column", type: "deleteTableColumn", blockId: "mixed-column-B" },
        {
          id: "insert-column",
          type: "insertTableColumn",
          blockId: "mixed-column-B",
          position: "before",
          cellTexts: ["New top", "New bottom"],
        },
      ],
      mode: "direct",
    });

    expect(result.skipped).toEqual([]);
    expect(view.state.doc.child(0).child(0).textContent).toBe("ANew top");
    expect(view.state.doc.child(0).child(1).textContent).toBe("CNew bottom");
  });

  test("maps column deletion through earlier operations in a mixed batch", () => {
    const rows = [
      ["A", "B"],
      ["C", "D"],
    ].map((texts) =>
      schema.node(
        "tableRow",
        null,
        texts.map((text) =>
          schema.node("tableCell", null, [
            schema.node("paragraph", { paraId: `mapped-column-${text}` }, [schema.text(text)]),
          ]),
        ),
      ),
    );
    const state = EditorState.create({
      schema,
      doc: schema.node("doc", null, [schema.node("table", null, rows)]),
    });
    const view = makeView(state);

    const result = applyFolioAIEditOperations({
      view,
      snapshot: createFolioAIEditSnapshot(state.doc),
      operations: [
        { id: "delete-column", type: "deleteTableColumn", blockId: "mapped-column-C" },
        {
          id: "insert-row",
          type: "insertTableRow",
          blockId: "mapped-column-C",
          cellTexts: ["E", "F"],
        },
        {
          id: "insert-before-table",
          type: "insertBeforeBlock",
          blockId: "mapped-column-A",
          text: "Before table",
        },
      ],
      mode: "direct",
    });

    expect(result.skipped).toEqual([]);
    expect(view.state.doc.child(0).textContent).toBe("Before table");
    const updatedTable = view.state.doc.child(1);
    expect(TableMap.get(updatedTable).width).toBe(1);
    expect(updatedTable.childCount).toBe(3);
    expect(updatedTable.textContent).toBe("BDF");
  });

  test("table-column deletes create one revision across every removed cell", () => {
    const makeTrackedColumnDeletionState = () => {
      const rows = [
        ["A", "B"],
        ["C", "D"],
      ].map((texts) =>
        schema.node(
          "tableRow",
          null,
          texts.map((text) =>
            schema.node("tableCell", null, [
              schema.node("paragraph", { paraId: `tracked-delete-column-${text}` }, [
                schema.text(text),
              ]),
            ]),
          ),
        ),
      );
      return EditorState.create({
        schema,
        doc: schema.node("doc", null, [schema.node("table", null, rows)]),
      });
    };

    const accepting = makeView(makeTrackedColumnDeletionState());
    const acceptedResult = applyFolioAIEditOperations({
      view: accepting,
      snapshot: createFolioAIEditSnapshot(accepting.state.doc),
      operations: [
        {
          id: "delete-column",
          type: "deleteTableColumn",
          blockId: "tracked-delete-column-B",
        },
      ],
      mode: "tracked-changes",
    });
    const revisionId = acceptedResult.applied.at(0)?.revisionId;
    if (revisionId === undefined) {
      throw new Error("expected a column deletion revision");
    }
    expect(acceptedResult).toEqual({
      applied: [{ id: "delete-column", revisionId, revisionIds: [revisionId] }],
      skipped: [],
    });
    const pendingTable = accepting.state.doc.child(0);
    expect(TableMap.get(pendingTable).width).toBe(2);
    expect(pendingTable.child(0).child(1).attrs["cellMarker"]).toEqual({
      kind: "del",
      info: { revisionId, author: "AI", date: expect.any(String) },
    });
    expect(pendingTable.child(1).child(1).attrs["cellMarker"]).toEqual({
      kind: "del",
      info: { revisionId, author: "AI", date: expect.any(String) },
    });
    expect(
      getTrackedChangesFromDoc(accepting.state.doc).filter(({ type }) => type === "cellDeleted"),
    ).toEqual([
      {
        id: revisionId,
        type: "cellDeleted",
        author: "AI",
        date: expect.any(String),
        text: "B\nD",
        blockId: expect.any(String),
      },
    ]);

    expect(acceptAIEditRevision(revisionId)(accepting.state, accepting.dispatch)).toBe(true);
    const acceptedTable = accepting.state.doc.child(0);
    expect(TableMap.get(acceptedTable).width).toBe(1);
    expect(acceptedTable.textContent).toBe("AC");

    const rejecting = makeView(makeTrackedColumnDeletionState());
    const rejectedResult = applyFolioAIEditOperations({
      view: rejecting,
      snapshot: createFolioAIEditSnapshot(rejecting.state.doc),
      operations: [
        {
          id: "delete-column",
          type: "deleteTableColumn",
          blockId: "tracked-delete-column-B",
        },
      ],
      mode: "tracked-changes",
    });
    const rejectedRevisionId = rejectedResult.applied.at(0)?.revisionId;
    if (rejectedRevisionId === undefined) {
      throw new Error("expected a column deletion revision");
    }
    expect(rejectAIEditRevision(rejectedRevisionId)(rejecting.state, rejecting.dispatch)).toBe(
      true,
    );
    const rejectedTable = rejecting.state.doc.child(0);
    expect(TableMap.get(rejectedTable).width).toBe(2);
    expect(rejectedTable.textContent).toBe("ABCD");
    expect(rejectedTable.child(0).child(1).attrs["cellMarker"]).toBeNull();
    expect(rejectedTable.child(1).child(1).attrs["cellMarker"]).toBeNull();
  });

  test("tracked column deletion rejects a column crossed by a colspan", () => {
    const table = schema.node("table", null, [
      schema.node("tableRow", null, [
        schema.node("tableCell", { colspan: 2 }, [
          schema.node("paragraph", { paraId: "tracked-delete-column-span" }, [schema.text("A")]),
        ]),
      ]),
      schema.node("tableRow", null, [
        schema.node("tableCell", null, [
          schema.node("paragraph", { paraId: "tracked-delete-column-left" }, [schema.text("B")]),
        ]),
        schema.node("tableCell", null, [
          schema.node("paragraph", { paraId: "tracked-delete-column-right" }, [schema.text("C")]),
        ]),
      ]),
    ]);
    const state = EditorState.create({ schema, doc: schema.node("doc", null, [table]) });
    const view = makeView(state);

    const result = applyFolioAIEditOperations({
      view,
      snapshot: createFolioAIEditSnapshot(state.doc),
      operations: [
        {
          id: "delete-column",
          type: "deleteTableColumn",
          blockId: "tracked-delete-column-right",
        },
      ],
      mode: "tracked-changes",
    });

    expect(result).toEqual({
      applied: [],
      skipped: [{ id: "delete-column", reason: "unsupportedBlock" }],
    });
    expect(view.state.doc).toEqual(state.doc);
  });

  test("table-column inserts create one revision across every new cell", () => {
    const makeTrackedColumnState = () => {
      const rows = [
        ["A", "B"],
        ["C", "D"],
      ].map((texts) =>
        schema.node(
          "tableRow",
          null,
          texts.map((text) =>
            schema.node("tableCell", null, [
              schema.node("paragraph", { paraId: `tracked-column-${text}` }, [schema.text(text)]),
            ]),
          ),
        ),
      );
      return EditorState.create({
        schema,
        doc: schema.node("doc", null, [schema.node("table", null, rows)]),
      });
    };

    const accepting = makeView(makeTrackedColumnState());
    const acceptedResult = applyFolioAIEditOperations({
      view: accepting,
      snapshot: createFolioAIEditSnapshot(accepting.state.doc),
      operations: [
        {
          id: "insert-column",
          type: "insertTableColumn",
          blockId: "tracked-column-A",
          cellTexts: ["New top", "New bottom"],
        },
      ],
      mode: "tracked-changes",
    });
    const revisionId = acceptedResult.applied.at(0)?.revisionId;
    if (revisionId === undefined) {
      throw new Error("expected a column insertion revision");
    }
    expect(acceptedResult).toEqual({
      applied: [{ id: "insert-column", revisionId, revisionIds: [revisionId] }],
      skipped: [],
    });
    const pendingTable = accepting.state.doc.child(0);
    expect(pendingTable.child(0).child(1).attrs["cellMarker"]).toEqual({
      kind: "ins",
      info: { revisionId, author: "AI", date: expect.any(String) },
    });
    expect(pendingTable.child(1).child(1).attrs["cellMarker"]).toEqual({
      kind: "ins",
      info: { revisionId, author: "AI", date: expect.any(String) },
    });
    expect(
      getTrackedChangesFromDoc(accepting.state.doc).filter(({ type }) => type === "cellInserted"),
    ).toEqual([
      {
        id: revisionId,
        type: "cellInserted",
        author: "AI",
        date: expect.any(String),
        text: "New top\nNew bottom",
        blockId: expect.any(String),
      },
    ]);

    expect(acceptAIEditRevision(revisionId)(accepting.state, accepting.dispatch)).toBe(true);
    const acceptedTable = accepting.state.doc.child(0);
    expect(TableMap.get(acceptedTable).width).toBe(3);
    expect(acceptedTable.child(0).child(1).attrs["cellMarker"]).toBeNull();
    expect(acceptedTable.child(1).child(1).attrs["cellMarker"]).toBeNull();

    const rejecting = makeView(makeTrackedColumnState());
    const rejectedResult = applyFolioAIEditOperations({
      view: rejecting,
      snapshot: createFolioAIEditSnapshot(rejecting.state.doc),
      operations: [
        {
          id: "insert-column",
          type: "insertTableColumn",
          blockId: "tracked-column-A",
          cellTexts: ["New top", "New bottom"],
        },
      ],
      mode: "tracked-changes",
    });
    const rejectedRevisionId = rejectedResult.applied.at(0)?.revisionId;
    if (rejectedRevisionId === undefined) {
      throw new Error("expected a column insertion revision");
    }
    expect(accepting.state.doc.textContent).toBe("ANew topBCNew bottomD");
    expect(rejecting.state.doc.textContent).toBe("ANew topBCNew bottomD");
    expect(rejectAIEditRevision(rejectedRevisionId)(rejecting.state, rejecting.dispatch)).toBe(
      true,
    );
    const rejectedTable = rejecting.state.doc.child(0);
    expect(TableMap.get(rejectedTable).width).toBe(2);
    expect(rejectedTable.textContent).toBe("ABCD");
  });

  test("tracked column insertion rejects a boundary crossed by a colspan", () => {
    const table = schema.node("table", null, [
      schema.node("tableRow", null, [
        schema.node("tableCell", { colspan: 2 }, [
          schema.node("paragraph", { paraId: "tracked-column-span" }, [schema.text("A")]),
        ]),
      ]),
      schema.node("tableRow", null, [
        schema.node("tableCell", null, [
          schema.node("paragraph", { paraId: "tracked-column-left" }, [schema.text("B")]),
        ]),
        schema.node("tableCell", null, [
          schema.node("paragraph", { paraId: "tracked-column-right" }, [schema.text("C")]),
        ]),
      ]),
    ]);
    const state = EditorState.create({ schema, doc: schema.node("doc", null, [table]) });
    const view = makeView(state);

    const result = applyFolioAIEditOperations({
      view,
      snapshot: createFolioAIEditSnapshot(state.doc),
      operations: [
        {
          id: "insert-column",
          type: "insertTableColumn",
          blockId: "tracked-column-left",
        },
      ],
      mode: "tracked-changes",
    });

    expect(result).toEqual({
      applied: [],
      skipped: [{ id: "insert-column", reason: "unsupportedBlock" }],
    });
    expect(view.state.doc).toEqual(state.doc);
  });

  test("merges a rectangular table region and preserves cell content order", () => {
    const rows = [
      ["A", "B"],
      ["C", "D"],
    ].map((texts) =>
      schema.node(
        "tableRow",
        null,
        texts.map((text) =>
          schema.node("tableCell", null, [
            schema.node("paragraph", { paraId: `merge-${text}` }, [schema.text(text)]),
          ]),
        ),
      ),
    );
    const state = EditorState.create({
      schema,
      doc: schema.node("doc", null, [schema.node("table", null, rows)]),
    });
    const view = makeView(state);

    const result = applyFolioAIEditOperations({
      view,
      snapshot: createFolioAIEditSnapshot(state.doc),
      operations: [
        {
          id: "merge-cells",
          type: "mergeTableCells",
          blockId: "merge-A",
          endBlockId: "merge-D",
        },
      ],
      mode: "direct",
    });

    expect(result).toEqual({ applied: [{ id: "merge-cells" }], skipped: [] });
    const table = view.state.doc.child(0);
    expect(TableMap.get(table)).toMatchObject({ width: 2, height: 2 });
    expect(table.child(0).childCount).toBe(1);
    expect(table.child(1).childCount).toBe(0);
    expect(table.child(0).child(0).attrs).toMatchObject({ colspan: 2, rowspan: 2 });
    expect(table.child(0).child(0).textContent).toBe("ABCD");
    expect(table.child(0).child(0).childCount).toBe(4);
  });

  test("merges a region that fully encloses an existing span", () => {
    const table = schema.node("table", null, [
      schema.node("tableRow", null, [
        schema.node("tableCell", { colspan: 2 }, [
          schema.node("paragraph", { paraId: "merge-span-A" }, [schema.text("A")]),
        ]),
      ]),
      schema.node("tableRow", null, [
        schema.node("tableCell", null, [
          schema.node("paragraph", { paraId: "merge-span-B" }, [schema.text("B")]),
        ]),
        schema.node("tableCell", null, [
          schema.node("paragraph", { paraId: "merge-span-C" }, [schema.text("C")]),
        ]),
      ]),
    ]);
    const state = EditorState.create({ schema, doc: schema.node("doc", null, [table]) });
    const view = makeView(state);

    const result = applyFolioAIEditOperations({
      view,
      snapshot: createFolioAIEditSnapshot(state.doc),
      operations: [
        {
          id: "merge-span",
          type: "mergeTableCells",
          blockId: "merge-span-A",
          endBlockId: "merge-span-C",
        },
      ],
      mode: "direct",
    });

    expect(result.skipped).toEqual([]);
    const merged = view.state.doc.child(0).child(0).child(0);
    expect(merged.attrs).toMatchObject({ colspan: 2, rowspan: 2 });
    expect(merged.textContent).toBe("ABC");
  });

  test("rejects a merge that cuts through an existing span without mutation", () => {
    const table = schema.node("table", null, [
      schema.node("tableRow", null, [
        schema.node("tableCell", { colspan: 2 }, [
          schema.node("paragraph", { paraId: "partial-A" }, [schema.text("A")]),
        ]),
        schema.node("tableCell", null, [
          schema.node("paragraph", { paraId: "partial-D" }, [schema.text("D")]),
        ]),
      ]),
      schema.node("tableRow", null, [
        schema.node("tableCell", null, [
          schema.node("paragraph", { paraId: "partial-B" }, [schema.text("B")]),
        ]),
        schema.node("tableCell", null, [
          schema.node("paragraph", { paraId: "partial-C" }, [schema.text("C")]),
        ]),
        schema.node("tableCell", null, [
          schema.node("paragraph", { paraId: "partial-E" }, [schema.text("E")]),
        ]),
      ]),
    ]);
    const state = EditorState.create({ schema, doc: schema.node("doc", null, [table]) });
    const view = makeView(state);

    const result = applyFolioAIEditOperations({
      view,
      snapshot: createFolioAIEditSnapshot(state.doc),
      operations: [
        {
          id: "partial-merge",
          type: "mergeTableCells",
          blockId: "partial-D",
          endBlockId: "partial-C",
        },
      ],
      mode: "direct",
    });

    expect(result).toEqual({
      applied: [],
      skipped: [{ id: "partial-merge", reason: "unsupportedBlock" }],
    });
    expect(view.state.doc).toEqual(state.doc);
  });

  test("applies cell text edits before merging and maps the table through earlier insertions", () => {
    const table = schema.node("table", null, [
      schema.node("tableRow", null, [
        schema.node("tableCell", null, [
          schema.node("paragraph", { paraId: "mapped-merge-A" }, [schema.text("A")]),
        ]),
        schema.node("tableCell", null, [
          schema.node("paragraph", { paraId: "mapped-merge-B" }, [schema.text("B")]),
        ]),
      ]),
    ]);
    const state = EditorState.create({ schema, doc: schema.node("doc", null, [table]) });
    const view = makeView(state);

    const result = applyFolioAIEditOperations({
      view,
      snapshot: createFolioAIEditSnapshot(state.doc),
      operations: [
        {
          id: "merge",
          type: "mergeTableCells",
          blockId: "mapped-merge-A",
          endBlockId: "mapped-merge-B",
        },
        {
          id: "replace",
          type: "replaceInBlock",
          blockId: "mapped-merge-B",
          find: "B",
          replace: "Changed",
        },
        {
          id: "insert-before",
          type: "insertBeforeBlock",
          blockId: "mapped-merge-A",
          text: "Before",
        },
      ],
      mode: "direct",
    });

    expect(result.skipped).toEqual([]);
    expect(view.state.doc.child(0).textContent).toBe("Before");
    const merged = view.state.doc.child(1).child(0).child(0);
    expect(merged.attrs["colspan"]).toBe(2);
    expect(merged.textContent).toBe("AChanged");
  });

  test("rejects conflicting or duplicate merge claims deterministically", () => {
    const rows = [
      ["A", "B", "C"],
      ["D", "E", "F"],
    ].map((texts) =>
      schema.node(
        "tableRow",
        null,
        texts.map((text) =>
          schema.node("tableCell", null, [
            schema.node("paragraph", { paraId: `claim-${text}` }, [schema.text(text)]),
          ]),
        ),
      ),
    );
    const state = EditorState.create({
      schema,
      doc: schema.node("doc", null, [schema.node("table", null, rows)]),
    });
    const view = makeView(state);

    const result = applyFolioAIEditOperations({
      view,
      snapshot: createFolioAIEditSnapshot(state.doc),
      operations: [
        {
          id: "first",
          type: "mergeTableCells",
          blockId: "claim-A",
          endBlockId: "claim-B",
        },
        {
          id: "duplicate",
          type: "mergeTableCells",
          blockId: "claim-B",
          endBlockId: "claim-A",
        },
        {
          id: "overlap",
          type: "mergeTableCells",
          blockId: "claim-B",
          endBlockId: "claim-C",
        },
      ],
      mode: "direct",
    });

    expect(result.applied).toEqual([{ id: "first" }]);
    expect(result.skipped).toEqual([
      { id: "duplicate", reason: "noopOperation" },
      { id: "overlap", reason: "unsupportedBlock" },
    ]);
    expect(view.state.doc.child(0).child(0).child(0).textContent).toBe("AB");
  });

  test("skips a merge when the same batch changes that table's structure", () => {
    const table = schema.node("table", null, [
      schema.node("tableRow", null, [
        schema.node("tableCell", null, [
          schema.node("paragraph", { paraId: "structural-A" }, [schema.text("A")]),
        ]),
        schema.node("tableCell", null, [
          schema.node("paragraph", { paraId: "structural-B" }, [schema.text("B")]),
        ]),
      ]),
    ]);
    const state = EditorState.create({ schema, doc: schema.node("doc", null, [table]) });
    const view = makeView(state);

    const result = applyFolioAIEditOperations({
      view,
      snapshot: createFolioAIEditSnapshot(state.doc),
      operations: [
        {
          id: "merge",
          type: "mergeTableCells",
          blockId: "structural-A",
          endBlockId: "structural-B",
        },
        {
          id: "insert-column",
          type: "insertTableColumn",
          blockId: "structural-B",
        },
      ],
      mode: "direct",
    });

    expect(result.applied).toEqual([{ id: "insert-column" }]);
    expect(result.skipped).toEqual([{ id: "merge", reason: "unsupportedBlock" }]);
    expect(TableMap.get(view.state.doc.child(0)).width).toBe(3);
  });

  test("tracks an empty vertical merge as one reversible cell revision", () => {
    const table = schema.node("table", null, [
      schema.node("tableRow", null, [
        schema.node("tableCell", null, [
          schema.node("paragraph", { paraId: "tracked-merge-start" }, [schema.text("A")]),
        ]),
        schema.node("tableCell", null, [
          schema.node("paragraph", { paraId: "tracked-merge-side-a" }, [schema.text("X")]),
        ]),
      ]),
      schema.node("tableRow", null, [
        schema.node(
          "tableCell",
          { verticalAlign: "bottom", _originalFormatting: { verticalAlign: "bottom" } },
          [schema.node("paragraph", { paraId: "tracked-merge-empty-b" })],
        ),
        schema.node("tableCell", null, [
          schema.node("paragraph", { paraId: "tracked-merge-side-b" }, [schema.text("Y")]),
        ]),
      ]),
      schema.node("tableRow", null, [
        schema.node(
          "tableCell",
          { verticalAlign: "center", _originalFormatting: { verticalAlign: "center" } },
          [schema.node("paragraph", { paraId: "tracked-merge-empty-c" })],
        ),
        schema.node("tableCell", null, [
          schema.node("paragraph", { paraId: "tracked-merge-side-c" }, [schema.text("Z")]),
        ]),
      ]),
    ]);
    const state = EditorState.create({ schema, doc: schema.node("doc", null, [table]) });
    const view = makeView(state);

    const result = applyFolioAIEditOperations({
      view,
      snapshot: createFolioAIEditSnapshot(state.doc),
      operations: [
        {
          id: "tracked-merge",
          type: "mergeTableCells",
          blockId: "tracked-merge-start",
          rowCount: 3,
        },
      ],
      mode: "tracked-changes",
    });

    expect(result.skipped).toEqual([]);
    const revisionId = result.applied.at(0)?.revisionId;
    expect(typeof revisionId).toBe("number");
    expect(result.applied.at(0)?.revisionIds).toEqual([revisionId]);
    const mergedCell = view.state.doc.child(0).child(0).child(0);
    expect(mergedCell.attrs["rowspan"]).toBe(3);
    expect(view.state.doc.child(0).child(1).childCount).toBe(1);
    expect(view.state.doc.child(0).child(2).childCount).toBe(1);
    expect(mergedCell.attrs["_docxVMergeContinuationCells"]).toEqual([
      expect.objectContaining({
        formatting: expect.objectContaining({ vMerge: "continue", verticalAlign: "bottom" }),
        structuralChange: {
          type: "tableCellMerge",
          info: { id: revisionId, author: "AI", date: expect.any(String) },
          verticalMerge: "continue",
          verticalMergeOriginal: "rest",
        },
      }),
      expect.objectContaining({
        formatting: expect.objectContaining({ vMerge: "continue", verticalAlign: "center" }),
        structuralChange: expect.objectContaining({
          type: "tableCellMerge",
          info: expect.objectContaining({ id: revisionId }),
        }),
      }),
    ]);

    if (revisionId === undefined) {
      throw new Error("expected a vertical merge revision");
    }
    expect(acceptAIEditRevision(revisionId)(view.state, view.dispatch)).toBe(true);
    const acceptedContinuations = view.state.doc.child(0).child(0).child(0).attrs[
      "_docxVMergeContinuationCells"
    ];
    expect(acceptedContinuations).toEqual([
      expect.not.objectContaining({ structuralChange: expect.anything() }),
      expect.not.objectContaining({ structuralChange: expect.anything() }),
    ]);

    const rejectingView = makeView(state);
    const rejectingResult = applyFolioAIEditOperations({
      view: rejectingView,
      snapshot: createFolioAIEditSnapshot(state.doc),
      operations: [
        {
          id: "tracked-merge",
          type: "mergeTableCells",
          blockId: "tracked-merge-start",
          rowCount: 3,
        },
      ],
      mode: "tracked-changes",
    });
    const rejectingRevisionId = rejectingResult.applied.at(0)?.revisionId;
    if (rejectingRevisionId === undefined) {
      throw new Error("expected a vertical merge revision");
    }
    expect(
      rejectAIEditRevision(rejectingRevisionId)(rejectingView.state, rejectingView.dispatch),
    ).toBe(true);
    const rejectedTable = rejectingView.state.doc.child(0);
    expect(rejectedTable.child(0).child(0).attrs["rowspan"]).toBe(1);
    expect(rejectedTable.child(1).child(0).attrs).toMatchObject({
      cellMarker: null,
      _originalFormatting: { verticalAlign: "bottom" },
    });
    expect(rejectedTable.child(2).child(0).attrs).toMatchObject({
      cellMarker: null,
      _originalFormatting: { verticalAlign: "center" },
    });
  });

  test("tracked merging rejects ambiguous topology without mutation", () => {
    const cases = [
      {
        name: "horizontal",
        table: schema.node("table", null, [
          schema.node("tableRow", null, [
            schema.node("tableCell", null, [
              schema.node("paragraph", { paraId: "tracked-merge-horizontal-a" }, [
                schema.text("A"),
              ]),
            ]),
            schema.node("tableCell", null, [
              schema.node("paragraph", { paraId: "tracked-merge-horizontal-b" }, [
                schema.text("B"),
              ]),
            ]),
          ]),
        ]),
        operation: {
          id: "tracked-merge-horizontal",
          type: "mergeTableCells" as const,
          blockId: "tracked-merge-horizontal-a",
          endBlockId: "tracked-merge-horizontal-b",
        },
      },
      {
        name: "content-bearing continuation",
        table: schema.node("table", null, [
          schema.node("tableRow", null, [
            schema.node("tableCell", null, [
              schema.node("paragraph", { paraId: "tracked-merge-content-a" }, [schema.text("A")]),
            ]),
          ]),
          schema.node("tableRow", null, [
            schema.node("tableCell", null, [
              schema.node("paragraph", { paraId: "tracked-merge-content-b" }, [schema.text("B")]),
            ]),
          ]),
        ]),
        operation: {
          id: "tracked-merge-content",
          type: "mergeTableCells" as const,
          blockId: "tracked-merge-content-a",
          endBlockId: "tracked-merge-content-b",
        },
      },
      {
        name: "marked continuation",
        table: schema.node("table", null, [
          schema.node("tableRow", null, [
            schema.node("tableCell", null, [
              schema.node("paragraph", { paraId: "tracked-merge-marked-a" }, [schema.text("A")]),
            ]),
          ]),
          schema.node("tableRow", null, [
            schema.node(
              "tableCell",
              {
                cellMarker: {
                  kind: "ins",
                  info: {
                    revisionId: 7,
                    author: "Existing",
                    date: "2026-07-17T00:00:00.000Z",
                  },
                },
              },
              [schema.node("paragraph", { paraId: "tracked-merge-marked-b" })],
            ),
          ]),
        ]),
        operation: {
          id: "tracked-merge-marked",
          type: "mergeTableCells" as const,
          blockId: "tracked-merge-marked-a",
          rowCount: 2,
        },
      },
    ];

    for (const { name, table, operation } of cases) {
      const state = EditorState.create({ schema, doc: schema.node("doc", null, [table]) });
      const view = makeView(state);
      expect(
        applyFolioAIEditOperations({
          view,
          snapshot: createFolioAIEditSnapshot(state.doc),
          operations: [operation],
          mode: "tracked-changes",
        }),
        name,
      ).toEqual({
        applied: [],
        skipped: [{ id: operation.id, reason: "unsupportedBlock" }],
      });
      expect(view.state.doc.eq(state.doc), name).toBe(true);
    }
  });

  test("treats a same-cell merge as a no-op in either mode", () => {
    const table = schema.node("table", null, [
      schema.node("tableRow", null, [
        schema.node("tableCell", null, [
          schema.node("paragraph", { paraId: "single-merge" }, [schema.text("A")]),
        ]),
      ]),
    ]);
    const state = EditorState.create({ schema, doc: schema.node("doc", null, [table]) });

    const directView = makeView(state);
    expect(
      applyFolioAIEditOperations({
        view: directView,
        snapshot: createFolioAIEditSnapshot(state.doc),
        operations: [
          {
            id: "same-cell",
            type: "mergeTableCells",
            blockId: "single-merge",
            endBlockId: "single-merge",
          },
        ],
        mode: "direct",
      }),
    ).toEqual({
      applied: [],
      skipped: [{ id: "same-cell", reason: "noopOperation" }],
    });

    const trackedView = makeView(state);
    expect(
      applyFolioAIEditOperations({
        view: trackedView,
        snapshot: createFolioAIEditSnapshot(state.doc),
        operations: [
          {
            id: "tracked-merge",
            type: "mergeTableCells",
            blockId: "single-merge",
            endBlockId: "single-merge",
          },
        ],
        mode: "tracked-changes",
      }),
    ).toEqual({
      applied: [],
      skipped: [{ id: "tracked-merge", reason: "noopOperation" }],
    });
  });

  test("splits a rectangular cell while preserving content and column widths", () => {
    const table = schema.node("table", null, [
      schema.node("tableRow", null, [
        schema.node("tableCell", { colspan: 2, rowspan: 2, colwidth: [100, 200] }, [
          schema.node("paragraph", { paraId: "split-all" }, [schema.text("Content")]),
        ]),
      ]),
      schema.node("tableRow"),
    ]);
    const state = EditorState.create({ schema, doc: schema.node("doc", null, [table]) });
    const view = makeView(state);

    const result = applyFolioAIEditOperations({
      view,
      snapshot: createFolioAIEditSnapshot(state.doc),
      operations: [{ id: "split-cell", type: "splitTableCell", blockId: "split-all" }],
      mode: "direct",
    });

    expect(result).toEqual({ applied: [{ id: "split-cell" }], skipped: [] });
    const updatedTable = view.state.doc.child(0);
    expect(TableMap.get(updatedTable)).toMatchObject({ width: 2, height: 2 });
    expect(updatedTable.child(0).childCount).toBe(2);
    expect(updatedTable.child(1).childCount).toBe(2);
    expect(updatedTable.child(0).child(0).textContent).toBe("Content");
    expect(updatedTable.child(0).child(1).textContent).toBe("");
    expect(updatedTable.child(1).textContent).toBe("");
    expect(updatedTable.child(0).child(0).attrs).toMatchObject({
      colspan: 1,
      rowspan: 1,
      colwidth: [100],
    });
    expect(updatedTable.child(0).child(1).attrs["colwidth"]).toEqual([200]);
    expect(updatedTable.child(1).child(0).attrs["colwidth"]).toEqual([100]);
    expect(updatedTable.child(1).child(1).attrs["colwidth"]).toEqual([200]);
  });

  test("tracks a vertical split as one reversible cell merge revision", () => {
    const continuationCells = [
      {
        type: "tableCell" as const,
        formatting: {
          vMerge: "continue" as const,
          verticalAlign: "bottom" as const,
        },
        content: [
          {
            type: "paragraph" as const,
            content: [
              {
                type: "run" as const,
                content: [{ type: "text" as const, text: "Second" }],
              },
            ],
          },
        ],
      },
      {
        type: "tableCell" as const,
        formatting: {
          vMerge: "continue" as const,
          verticalAlign: "center" as const,
        },
        content: [
          {
            type: "paragraph" as const,
            content: [
              {
                type: "run" as const,
                content: [{ type: "text" as const, text: "Third" }],
              },
            ],
          },
        ],
      },
    ];
    const table = schema.node("table", null, [
      schema.node("tableRow", null, [
        schema.node(
          "tableCell",
          {
            rowspan: 3,
            _originalFormatting: { vMerge: "restart" },
            _docxVMergeContinuationCells: continuationCells,
          },
          [schema.node("paragraph", { paraId: "tracked-split" }, [schema.text("Content")])],
        ),
      ]),
      schema.node("tableRow"),
      schema.node("tableRow"),
    ]);
    const state = EditorState.create({ schema, doc: schema.node("doc", null, [table]) });
    const view = makeView(state);

    const result = applyFolioAIEditOperations({
      view,
      snapshot: createFolioAIEditSnapshot(state.doc),
      operations: [{ id: "split-cell", type: "splitTableCell", blockId: "tracked-split" }],
      mode: "tracked-changes",
    });

    expect(result.skipped).toEqual([]);
    const revisionId = result.applied.at(0)?.revisionId;
    expect(typeof revisionId).toBe("number");
    expect(result.applied.at(0)?.revisionIds).toEqual([revisionId]);
    const splitTable = view.state.doc.child(0);
    expect(splitTable.child(0).child(0).attrs["rowspan"]).toBe(1);
    expect(splitTable.child(1).childCount).toBe(1);
    expect(splitTable.child(2).childCount).toBe(1);
    expect(splitTable.child(1).child(0).textContent).toBe("Second");
    expect(splitTable.child(2).child(0).textContent).toBe("Third");
    expect(splitTable.child(1).child(0).attrs).toMatchObject({
      cellMarker: {
        kind: "merge",
        info: {
          revisionId,
          author: "AI",
          date: expect.any(String),
        },
        verticalMergeOriginal: "continue",
      },
      _originalFormatting: { verticalAlign: "bottom" },
    });
    expect(splitTable.child(2).child(0).attrs).toMatchObject({
      cellMarker: {
        kind: "merge",
        info: { revisionId },
        verticalMergeOriginal: "continue",
      },
      _originalFormatting: { verticalAlign: "center" },
    });

    if (revisionId === undefined) {
      throw new Error("expected a vertical split revision");
    }
    expect(acceptAIEditRevision(revisionId)(view.state, view.dispatch)).toBe(true);
    expect(view.state.doc.child(0).child(1).child(0).attrs["cellMarker"]).toBeNull();
    expect(view.state.doc.child(0).child(2).child(0).attrs["cellMarker"]).toBeNull();
    expect(view.state.doc.child(0).child(1).child(0).textContent).toBe("Second");
    expect(view.state.doc.child(0).child(2).child(0).textContent).toBe("Third");

    const rejectingView = makeView(state);
    const rejectingResult = applyFolioAIEditOperations({
      view: rejectingView,
      snapshot: createFolioAIEditSnapshot(state.doc),
      operations: [{ id: "split-cell", type: "splitTableCell", blockId: "tracked-split" }],
      mode: "tracked-changes",
    });
    const rejectingRevisionId = rejectingResult.applied.at(0)?.revisionId;
    if (rejectingRevisionId === undefined) {
      throw new Error("expected a vertical split revision");
    }
    expect(
      rejectAIEditRevision(rejectingRevisionId)(rejectingView.state, rejectingView.dispatch),
    ).toBe(true);
    const restoredTable = rejectingView.state.doc.child(0);
    expect(restoredTable.child(0).child(0).attrs["rowspan"]).toBe(3);
    expect(restoredTable.child(1).childCount).toBe(0);
    expect(restoredTable.child(2).childCount).toBe(0);
    expect(restoredTable.child(0).child(0).attrs["_docxVMergeContinuationCells"]).toEqual(
      continuationCells,
    );
  });

  test("tracked splitting refuses to restore a stored continuation cell with gridSpan > 1", () => {
    // Regression guard: splitTrackedVerticalTableCell only ever restores a
    // single-column rectangle (the caller enforces
    // `rectangle.right - rectangle.left === 1`), so a restored continuation
    // cell must be colspan 1. `_docxVMergeContinuationCells` entries are
    // captured from a prior (possibly attacker-crafted) parse; a stored
    // gridSpan > 1 must be refused instead of silently widening the
    // restored cell across multiple columns.
    const continuationCells = [
      {
        type: "tableCell" as const,
        formatting: {
          vMerge: "continue" as const,
          gridSpan: 5,
        },
        content: [
          {
            type: "paragraph" as const,
            content: [{ type: "run" as const, content: [{ type: "text" as const, text: "Wide" }] }],
          },
        ],
      },
    ];
    const table = schema.node("table", null, [
      schema.node("tableRow", null, [
        schema.node(
          "tableCell",
          {
            rowspan: 2,
            _originalFormatting: { vMerge: "restart" },
            _docxVMergeContinuationCells: continuationCells,
          },
          [
            schema.node("paragraph", { paraId: "wide-continuation-split" }, [
              schema.text("Content"),
            ]),
          ],
        ),
      ]),
      schema.node("tableRow"),
    ]);
    const state = EditorState.create({ schema, doc: schema.node("doc", null, [table]) });
    const view = makeView(state);

    const result = applyFolioAIEditOperations({
      view,
      snapshot: createFolioAIEditSnapshot(state.doc),
      operations: [
        { id: "split-cell", type: "splitTableCell", blockId: "wide-continuation-split" },
      ],
      mode: "tracked-changes",
    });

    expect(result.applied).toEqual([]);
    expect(result.skipped).toEqual([{ id: "split-cell", reason: "unsupportedBlock" }]);
    const untouchedTable = view.state.doc.child(0);
    expect(untouchedTable.child(0).child(0).attrs["rowspan"]).toBe(2);
    expect(untouchedTable.child(1).childCount).toBe(0);
  });

  test("tracked splitting rejects a horizontal span without mutation", () => {
    const table = schema.node("table", null, [
      schema.node("tableRow", null, [
        schema.node("tableCell", { colspan: 2 }, [
          schema.node("paragraph", { paraId: "tracked-horizontal-split" }, [
            schema.text("Content"),
          ]),
        ]),
      ]),
    ]);
    const state = EditorState.create({ schema, doc: schema.node("doc", null, [table]) });
    const view = makeView(state);

    const result = applyFolioAIEditOperations({
      view,
      snapshot: createFolioAIEditSnapshot(state.doc),
      operations: [
        {
          id: "split-cell",
          type: "splitTableCell",
          blockId: "tracked-horizontal-split",
        },
      ],
      mode: "tracked-changes",
    });

    expect(result).toEqual({
      applied: [],
      skipped: [{ id: "split-cell", reason: "unsupportedBlock" }],
    });
    expect(view.state.doc.eq(state.doc)).toBe(true);
  });

  test("applies cell text edits before splitting and maps through an earlier insertion", () => {
    const table = schema.node("table", null, [
      schema.node("tableRow", null, [
        schema.node("tableCell", { colspan: 2 }, [
          schema.node("paragraph", { paraId: "mapped-split" }, [schema.text("Before")]),
        ]),
      ]),
    ]);
    const state = EditorState.create({ schema, doc: schema.node("doc", null, [table]) });
    const view = makeView(state);

    const result = applyFolioAIEditOperations({
      view,
      snapshot: createFolioAIEditSnapshot(state.doc),
      operations: [
        { id: "split", type: "splitTableCell", blockId: "mapped-split" },
        {
          id: "replace",
          type: "replaceInBlock",
          blockId: "mapped-split",
          find: "Before",
          replace: "Changed",
        },
        {
          id: "insert-before",
          type: "insertBeforeBlock",
          blockId: "mapped-split",
          text: "Outside",
        },
      ],
      mode: "direct",
    });

    expect(result.skipped).toEqual([]);
    expect(view.state.doc.child(0).textContent).toBe("Outside");
    const updatedRow = view.state.doc.child(1).child(0);
    expect(updatedRow.childCount).toBe(2);
    expect(updatedRow.child(0).textContent).toBe("Changed");
    expect(updatedRow.child(1).textContent).toBe("");
  });

  test("splits distinct cells once and rejects duplicate targets", () => {
    const table = schema.node("table", null, [
      schema.node("tableRow", null, [
        schema.node("tableCell", { colspan: 2 }, [
          schema.node("paragraph", { paraId: "split-left" }, [schema.text("A")]),
        ]),
        schema.node("tableCell", { colspan: 2 }, [
          schema.node("paragraph", { paraId: "split-right" }, [schema.text("B")]),
        ]),
      ]),
    ]);
    const state = EditorState.create({ schema, doc: schema.node("doc", null, [table]) });
    const view = makeView(state);

    const result = applyFolioAIEditOperations({
      view,
      snapshot: createFolioAIEditSnapshot(state.doc),
      operations: [
        { id: "left", type: "splitTableCell", blockId: "split-left" },
        { id: "left-again", type: "splitTableCell", blockId: "split-left" },
        { id: "right", type: "splitTableCell", blockId: "split-right" },
      ],
      mode: "direct",
    });

    expect(result.applied.map(({ id }) => id).toSorted()).toEqual(["left", "right"]);
    expect(result.skipped).toEqual([{ id: "left-again", reason: "noopOperation" }]);
    expect(view.state.doc.child(0).child(0).childCount).toBe(4);
    expect(view.state.doc.child(0).child(0).textContent).toBe("AB");
  });

  test("splits the nearest cell when the anchor is inside a nested table", () => {
    const innerTable = schema.node("table", null, [
      schema.node("tableRow", null, [
        schema.node("tableCell", { colspan: 2 }, [
          schema.node("paragraph", { paraId: "inner-split" }, [schema.text("Inner")]),
        ]),
      ]),
    ]);
    const outerTable = schema.node("table", null, [
      schema.node("tableRow", null, [
        schema.node("tableCell", null, [
          schema.node("paragraph", null, [schema.text("Outer")]),
          innerTable,
        ]),
      ]),
    ]);
    const state = EditorState.create({ schema, doc: schema.node("doc", null, [outerTable]) });
    const view = makeView(state);

    const result = applyFolioAIEditOperations({
      view,
      snapshot: createFolioAIEditSnapshot(state.doc),
      operations: [{ id: "split-inner", type: "splitTableCell", blockId: "inner-split" }],
      mode: "direct",
    });

    expect(result.skipped).toEqual([]);
    const updatedOuterTable = view.state.doc.child(0);
    expect(updatedOuterTable.child(0).childCount).toBe(1);
    const updatedInnerTable = updatedOuterTable.child(0).child(0).child(1);
    expect(updatedInnerTable.child(0).childCount).toBe(2);
    expect(updatedInnerTable.textContent).toBe("Inner");
  });

  test("rejects same-table split and merge operations as an ambiguous batch", () => {
    const table = schema.node("table", null, [
      schema.node("tableRow", null, [
        schema.node("tableCell", { colspan: 2 }, [
          schema.node("paragraph", { paraId: "shape-A" }, [schema.text("A")]),
        ]),
        schema.node("tableCell", null, [
          schema.node("paragraph", { paraId: "shape-B" }, [schema.text("B")]),
        ]),
      ]),
    ]);
    const state = EditorState.create({ schema, doc: schema.node("doc", null, [table]) });
    const view = makeView(state);

    const result = applyFolioAIEditOperations({
      view,
      snapshot: createFolioAIEditSnapshot(state.doc),
      operations: [
        { id: "split", type: "splitTableCell", blockId: "shape-A" },
        {
          id: "merge",
          type: "mergeTableCells",
          blockId: "shape-A",
          endBlockId: "shape-B",
        },
      ],
      mode: "direct",
    });

    expect(result.applied).toEqual([]);
    expect(result.skipped).toEqual([
      { id: "split", reason: "unsupportedBlock" },
      { id: "merge", reason: "unsupportedBlock" },
    ]);
    expect(view.state.doc).toEqual(state.doc);
  });

  test("treats an unspanned split as a no-op in both modes", () => {
    const table = schema.node("table", null, [
      schema.node("tableRow", null, [
        schema.node("tableCell", null, [
          schema.node("paragraph", { paraId: "plain-cell" }, [schema.text("A")]),
        ]),
      ]),
    ]);
    const state = EditorState.create({ schema, doc: schema.node("doc", null, [table]) });

    const directView = makeView(state);
    expect(
      applyFolioAIEditOperations({
        view: directView,
        snapshot: createFolioAIEditSnapshot(state.doc),
        operations: [{ id: "plain", type: "splitTableCell", blockId: "plain-cell" }],
        mode: "direct",
      }),
    ).toEqual({
      applied: [],
      skipped: [{ id: "plain", reason: "noopOperation" }],
    });

    const trackedView = makeView(state);
    expect(
      applyFolioAIEditOperations({
        view: trackedView,
        snapshot: createFolioAIEditSnapshot(state.doc),
        operations: [{ id: "tracked", type: "splitTableCell", blockId: "plain-cell" }],
        mode: "tracked-changes",
      }),
    ).toEqual({
      applied: [],
      skipped: [{ id: "tracked", reason: "noopOperation" }],
    });
  });

  test("table-row deletes create one structural revision in tracked-changes mode", () => {
    const table = schema.node("table", null, [
      schema.node("tableRow", null, [
        schema.node("tableCell", null, [
          schema.node("paragraph", { paraId: "tracked-delete" }, [schema.text("Cell")]),
        ]),
      ]),
      schema.node("tableRow", null, [
        schema.node("tableCell", null, [
          schema.node("paragraph", { paraId: "untouched-row" }, [schema.text("Untouched")]),
        ]),
      ]),
    ]);
    const state = EditorState.create({ schema, doc: schema.node("doc", null, [table]) });
    const view = makeView(state);

    const result = applyFolioAIEditOperations({
      view,
      snapshot: createFolioAIEditSnapshot(state.doc),
      operations: [{ id: "delete-row", type: "deleteTableRow", blockId: "tracked-delete" }],
      mode: "tracked-changes",
    });

    expect(result.skipped).toEqual([]);
    expect(result.applied).toHaveLength(1);
    const revisionId = result.applied.at(0)?.revisionId;
    expect(typeof revisionId).toBe("number");
    expect(result.applied.at(0)?.revisionIds).toEqual([revisionId]);
    const liveTable = view.state.doc.child(0);
    expect(liveTable.childCount).toBe(2);
    expect(liveTable.child(0).attrs["trDel"]).toEqual({
      revisionId,
      author: "AI",
      date: expect.any(String),
    });
    expect(liveTable.child(1).attrs["trDel"]).toBeNull();
  });

  test("tracked row deletion rejects a row with a pending structural revision", () => {
    const table = schema.node("table", null, [
      schema.node(
        "tableRow",
        {
          trIns: {
            revisionId: 10,
            author: "Reviewer",
            date: "2026-07-16",
          },
        },
        [
          schema.node("tableCell", null, [
            schema.node("paragraph", { paraId: "pending-row" }, [schema.text("Cell")]),
          ]),
        ],
      ),
    ]);
    const state = EditorState.create({ schema, doc: schema.node("doc", null, [table]) });
    const view = makeView(state);

    const result = applyFolioAIEditOperations({
      view,
      snapshot: createFolioAIEditSnapshot(state.doc),
      operations: [{ id: "delete-row", type: "deleteTableRow", blockId: "pending-row" }],
      mode: "tracked-changes",
    });

    expect(result).toEqual({
      applied: [],
      skipped: [{ id: "delete-row", reason: "unsupportedBlock" }],
    });
    expect(view.state.doc).toEqual(state.doc);
  });

  test("accepting a tracked row deletion repairs a vertical span", () => {
    const table = schema.node("table", null, [
      schema.node("tableRow", null, [
        schema.node("tableCell", { colspan: 2, rowspan: 2 }, [
          schema.node("paragraph", { paraId: "tracked-span-origin" }, [schema.text("A")]),
        ]),
        schema.node("tableCell", null, [
          schema.node("paragraph", { paraId: "tracked-span-peer" }, [schema.text("B")]),
        ]),
      ]),
      schema.node("tableRow", null, [
        schema.node("tableCell", null, [
          schema.node("paragraph", { paraId: "tracked-span-continuation" }, [schema.text("C")]),
        ]),
      ]),
    ]);
    const state = EditorState.create({ schema, doc: schema.node("doc", null, [table]) });
    const view = makeView(state);

    const result = applyFolioAIEditOperations({
      view,
      snapshot: createFolioAIEditSnapshot(state.doc),
      operations: [
        {
          id: "delete-row",
          type: "deleteTableRow",
          blockId: "tracked-span-origin",
        },
      ],
      mode: "tracked-changes",
    });
    const revisionId = result.applied.at(0)?.revisionId;
    if (revisionId === undefined) {
      throw new Error("expected a row deletion revision");
    }

    expect(acceptAIEditRevision(revisionId)(view.state, view.dispatch)).toBe(true);
    const acceptedTable = view.state.doc.child(0);
    expect(acceptedTable.childCount).toBe(1);
    expect(acceptedTable.child(0).childCount).toBe(2);
    expect(acceptedTable.child(0).child(0).attrs["colspan"]).toBe(2);
    expect(acceptedTable.child(0).child(0).attrs["rowspan"]).toBe(1);
    expect(acceptedTable.textContent).toBe("AC");
  });

  test("table-row inserts create one structural revision in tracked-changes mode", () => {
    const table = schema.node("table", null, [
      schema.node("tableRow", null, [
        schema.node("tableCell", null, [
          schema.node("paragraph", { paraId: "tracked-cell" }, [schema.text("Cell")]),
        ]),
      ]),
    ]);
    const state = EditorState.create({ schema, doc: schema.node("doc", null, [table]) });
    const snapshot = createFolioAIEditSnapshot(state.doc);
    const view = makeView(state);

    const result = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [{ id: "insert-row", type: "insertTableRow", blockId: "tracked-cell" }],
      mode: "tracked-changes",
    });

    expect(result.skipped).toEqual([]);
    expect(result.applied).toHaveLength(1);
    const revisionId = result.applied.at(0)?.revisionId;
    expect(typeof revisionId).toBe("number");
    expect(result.applied.at(0)?.revisionIds).toEqual([revisionId]);
    const liveTable = view.state.doc.child(0);
    expect(liveTable.childCount).toBe(2);
    expect(liveTable.child(1).attrs["trIns"]).toEqual({
      revisionId,
      author: "AI",
      date: expect.any(String),
    });
  });

  test("tracked row insertion rejects a boundary crossed by a vertical span", () => {
    const table = schema.node("table", null, [
      schema.node("tableRow", null, [
        schema.node("tableCell", { rowspan: 2 }, [
          schema.node("paragraph", { paraId: "spanned-cell" }, [schema.text("Cell")]),
        ]),
      ]),
      schema.node("tableRow"),
    ]);
    const state = EditorState.create({ schema, doc: schema.node("doc", null, [table]) });
    const view = makeView(state);

    const result = applyFolioAIEditOperations({
      view,
      snapshot: createFolioAIEditSnapshot(state.doc),
      operations: [{ id: "insert-row", type: "insertTableRow", blockId: "spanned-cell" }],
      mode: "tracked-changes",
    });

    expect(result).toEqual({
      applied: [],
      skipped: [{ id: "insert-row", reason: "unsupportedBlock" }],
    });
    expect(view.state.doc).toEqual(state.doc);
  });
});
