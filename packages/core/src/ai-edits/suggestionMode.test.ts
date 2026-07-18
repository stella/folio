/**
 * Suggested apply-mode + suggestion command tests.
 *
 * Covers:
 * - applying operations in `"suggested"` mode stamps provenance "suggested"
 *   and a suggestionId on the produced insertion/deletion marks;
 * - block/table operations report `unsupportedMode` under suggested mode;
 * - `getSuggestions` lists the pending suggestions;
 * - `acceptSuggestion` rewrites the marks to a normal (user) tracked change
 *   authored by the accepting user, so the change then serializes;
 * - `rejectSuggestion` inverse-applies (removes suggested-inserted text) and
 *   leaves nothing to serialize.
 */

import { describe, expect, test } from "bun:test";
import type { Mark } from "prosemirror-model";
import { EditorState } from "prosemirror-state";
import type { Transaction } from "prosemirror-state";

import {
  acceptSuggestion,
  getSuggestions,
  rejectSuggestion,
} from "../prosemirror/commands/comments";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { schema } from "../prosemirror/schema";
import { serializeParagraph } from "../docx/serializer/paragraphSerializer";
import type { Document, Paragraph } from "../types/document";
import { applyFolioAIEditOperations } from "./apply";
import { createFolioAIEditSnapshot, createFolioAITextRangeHandle } from "./snapshot";

const makeView = (text: string) => {
  const doc = schema.node("doc", null, [schema.node("paragraph", {}, [schema.text(text)])]);
  const view = {
    state: EditorState.create({ schema, doc }),
    dispatch(transaction: Transaction) {
      view.state = view.state.apply(transaction);
    },
  };
  return view;
};

const marksInDoc = (state: EditorState): Mark[] => {
  const marks: Mark[] = [];
  state.doc.descendants((node) => {
    if (node.isInline) {
      marks.push(...node.marks);
    }
    return undefined;
  });
  return marks;
};

const firstParagraph = (doc: Document): Paragraph => {
  const block = doc.package.document.content.at(0);
  if (block?.type !== "paragraph") {
    throw new Error("expected a paragraph block");
  }
  return block;
};

const applySuggestedReplace = (
  view: ReturnType<typeof makeView>,
  find: string,
  replace: string,
) => {
  const snapshot = createFolioAIEditSnapshot(view.state.doc);
  const block = snapshot.blocks.at(0);
  if (!block) {
    throw new Error("expected a block");
  }
  return applyFolioAIEditOperations({
    view,
    snapshot,
    operations: [{ id: "op-1", type: "replaceInBlock", blockId: block.id, find, replace }],
    mode: "suggested",
    author: "AI",
  });
};

describe("suggested apply mode", () => {
  test("stamps provenance and suggestionId on produced marks and reports them", () => {
    const view = makeView("the quick brown fox");
    const result = applySuggestedReplace(view, "quick", "swift");

    const applied = result.applied.at(0);
    expect(applied?.id).toBe("op-1");
    // Defaults suggestionId to the operation id when the caller omits one.
    expect(applied?.suggestionId).toBe("op-1");

    const trackedMarks = marksInDoc(view.state).filter(
      (mark) => mark.type.name === "insertion" || mark.type.name === "deletion",
    );
    expect(trackedMarks.length).toBeGreaterThan(0);
    for (const mark of trackedMarks) {
      expect(mark.attrs["provenance"]).toBe("suggested");
      expect(mark.attrs["suggestionId"]).toBe("op-1");
    }
  });

  test("honours a caller-supplied suggestionId", () => {
    const view = makeView("the quick brown fox");
    const snapshot = createFolioAIEditSnapshot(view.state.doc);
    const block = snapshot.blocks.at(0)!;
    const result = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [
        {
          id: "op-1",
          type: "replaceInBlock",
          blockId: block.id,
          find: "quick",
          replace: "swift",
          suggestionId: "group-A",
        },
      ],
      mode: "suggested",
      author: "AI",
    });
    expect(result.applied.at(0)?.suggestionId).toBe("group-A");
    expect(
      marksInDoc(view.state)
        .filter((m) => m.type.name === "insertion" || m.type.name === "deletion")
        .every((m) => m.attrs["suggestionId"] === "group-A"),
    ).toBe(true);
  });

  test("cell merge reports unsupportedMode under suggested mode", () => {
    // Real 2x1 table so the operation resolves and the skip provably comes
    // from the suggested-mode allowlist, not a failed table resolution.
    const cellNode = (text: string, attrs: Record<string, unknown> = {}) =>
      schema.nodes["tableCell"]!.create(attrs, [schema.node("paragraph", {}, [schema.text(text)])]);
    const doc = schema.node("doc", null, [
      schema.nodes["table"]!.create({}, [
        schema.nodes["tableRow"]!.create({}, [cellNode("top")]),
        schema.nodes["tableRow"]!.create({}, [cellNode("bottom")]),
      ]),
    ]);
    const view = {
      state: EditorState.create({ schema, doc }),
      dispatch(transaction: Transaction) {
        view.state = view.state.apply(transaction);
      },
    };
    const snapshot = createFolioAIEditSnapshot(view.state.doc);
    const anchor = snapshot.blocks.find((block) => block.text === "top");
    if (!anchor) {
      throw new Error("expected the first-cell paragraph in the snapshot");
    }
    const result = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [{ id: "merge", type: "mergeTableCells", blockId: anchor.id, rowCount: 2 }],
      mode: "suggested",
      author: "AI",
    });
    expect(result.applied).toEqual([]);
    expect(result.skipped).toEqual([{ id: "merge", reason: "unsupportedMode" }]);
  });

  test("cell split reports unsupportedMode under suggested mode", () => {
    // A rowspan-2 cell so splitTableCell resolves (a 1x1 cell would skip as a
    // no-op before reaching the suggested-mode allowlist).
    const cellNode = (text: string, attrs: Record<string, unknown> = {}) =>
      schema.nodes["tableCell"]!.create(attrs, [schema.node("paragraph", {}, [schema.text(text)])]);
    const doc = schema.node("doc", null, [
      schema.nodes["table"]!.create({}, [
        schema.nodes["tableRow"]!.create({}, [
          cellNode("merged", { rowspan: 2 }),
          cellNode("r1c2"),
        ]),
        schema.nodes["tableRow"]!.create({}, [cellNode("r2c2")]),
      ]),
    ]);
    const view = {
      state: EditorState.create({ schema, doc }),
      dispatch(transaction: Transaction) {
        view.state = view.state.apply(transaction);
      },
    };
    const snapshot = createFolioAIEditSnapshot(view.state.doc);
    const anchor = snapshot.blocks.find((block) => block.text === "merged");
    if (!anchor) {
      throw new Error("expected the merged-cell paragraph in the snapshot");
    }
    const result = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [{ id: "split", type: "splitTableCell", blockId: anchor.id }],
      mode: "suggested",
      author: "AI",
    });
    expect(result.applied).toEqual([]);
    expect(result.skipped).toEqual([{ id: "split", reason: "unsupportedMode" }]);
  });

  test("insertAfterBlock in suggested mode flags the block and strips it", () => {
    const view = makeView("anchor block");
    const snapshot = createFolioAIEditSnapshot(view.state.doc);
    const block = snapshot.blocks.at(0)!;
    const result = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [
        { id: "ins", type: "insertAfterBlock", blockId: block.id, text: "proposed new block" },
      ],
      mode: "suggested",
      author: "AI",
    });
    expect(result.applied.at(0)?.suggestionId).toBe("ins");

    let flagged = false;
    view.state.doc.descendants((node) => {
      if (node.attrs["_suggestedInsert"]) {
        flagged = true;
      }
      return undefined;
    });
    expect(flagged).toBe(true);

    const model = fromProseDoc(view.state.doc);
    expect(JSON.stringify(model.package.document.content)).not.toContain("proposed new block");
  });

  test("deleteBlock in suggested mode keeps the text after the serialization strip", () => {
    const view = makeView("delete me please");
    const snapshot = createFolioAIEditSnapshot(view.state.doc);
    const block = snapshot.blocks.at(0)!;
    const result = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [{ id: "del", type: "deleteBlock", blockId: block.id }],
      mode: "suggested",
      author: "AI",
    });
    expect(result.applied.at(0)?.suggestionId).toBe("del");

    const xml = serializeParagraph(firstParagraph(fromProseDoc(view.state.doc)));
    expect(xml).toContain("delete me please");
    expect(xml).not.toContain("<w:del");
  });

  test("suggested formatRange stamps runPropertyChange and reverts on serialize", () => {
    const view = makeView("format me please");
    const snapshot = createFolioAIEditSnapshot(view.state.doc);
    const block = snapshot.blocks.at(0)!;
    const range = createFolioAITextRangeHandle({
      blockId: block.id,
      text: block.text,
      startOffset: 0,
      endOffset: 6,
    });
    if (!range) {
      throw new Error("expected a range");
    }
    const result = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [{ id: "fmt", type: "formatRange", range, formatting: { bold: true } }],
      mode: "suggested",
      author: "AI",
    });
    expect(result.applied.at(0)?.suggestionId).toBe("fmt");

    const rpc = marksInDoc(view.state).find((m) => m.type.name === "runPropertyChange");
    expect(rpc?.attrs["provenance"]).toBe("suggested");
    expect(rpc?.attrs["suggestionId"]).toBe("fmt");
    // The live doc shows the proposed bold; the strip reverts it on serialize.
    expect(marksInDoc(view.state).some((m) => m.type.name === "bold")).toBe(true);

    const xml = serializeParagraph(firstParagraph(fromProseDoc(view.state.doc)));
    expect(xml).not.toContain("<w:rPrChange");
    expect(xml).not.toContain("<w:b/>");
  });

  test("suggested changes are stripped from serialized output until accepted", () => {
    const view = makeView("the quick brown fox");
    applySuggestedReplace(view, "quick", "swift");
    const paragraph = firstParagraph(fromProseDoc(view.state.doc));
    const xml = serializeParagraph(paragraph);
    expect(xml).not.toContain("swift");
    expect(xml).not.toContain("<w:ins");
    expect(xml).not.toContain("<w:del");
  });
});

describe("suggestion commands", () => {
  test("getSuggestions lists the pending suggestion with its kinds", () => {
    const view = makeView("the quick brown fox");
    applySuggestedReplace(view, "quick", "swift");

    const suggestions = getSuggestions(view.state);
    expect(suggestions.length).toBe(1);
    const suggestion = suggestions.at(0)!;
    expect(suggestion.suggestionId).toBe("op-1");
    expect(suggestion.kinds).toContain("insertion");
    expect(suggestion.kinds).toContain("deletion");
    expect(suggestion.ranges.length).toBeGreaterThan(0);
  });

  test("acceptSuggestion converts marks to a user tracked change that serializes", () => {
    const view = makeView("the quick brown fox");
    applySuggestedReplace(view, "quick", "swift");

    const accepted = acceptSuggestion("op-1", {
      author: "Alice",
      date: "2026-07-17T12:00:00.000Z",
    })(view.state, view.dispatch);
    expect(accepted).toBe(true);

    // No suggested provenance remains; author is now the accepting user.
    const tracked = marksInDoc(view.state).filter(
      (m) => m.type.name === "insertion" || m.type.name === "deletion",
    );
    expect(tracked.length).toBeGreaterThan(0);
    for (const mark of tracked) {
      expect(mark.attrs["provenance"]).toBe("user");
      expect(mark.attrs["suggestionId"]).toBeNull();
      expect(mark.attrs["author"]).toBe("Alice");
    }
    expect(getSuggestions(view.state)).toEqual([]);

    // The accepted change now serializes as a real tracked change.
    const xml = serializeParagraph(firstParagraph(fromProseDoc(view.state.doc)));
    expect(xml).toContain("<w:ins");
    expect(xml).toContain("swift");
    expect(xml).toContain('w:author="Alice"');
  });

  test("rejectSuggestion removes the proposed insertion text", () => {
    const view = makeView("the quick brown fox");
    applySuggestedReplace(view, "quick", "swift");

    const rejected = rejectSuggestion("op-1")(view.state, view.dispatch);
    expect(rejected).toBe(true);

    // The suggested replacement text is gone and the original survives.
    expect(view.state.doc.textContent).toBe("the quick brown fox");
    expect(getSuggestions(view.state)).toEqual([]);
    const xml = serializeParagraph(firstParagraph(fromProseDoc(view.state.doc)));
    expect(xml).not.toContain("swift");
    expect(xml).not.toContain("<w:ins");
    expect(xml).not.toContain("<w:del");
  });
});
