/**
 * Unit tests for ParaIdAllocatorExtension.
 *
 * The plugin is lifted from eigenpal upstream (Apache-2.0). These
 * tests cover the contract we rely on downstream: every paragraph
 * ends up with a non-empty `paraId` and no two paragraphs share one,
 * including after paste-style Slice insertions that import a
 * paragraph carrying an id that already exists in the doc.
 *
 * The "Id lifecycle rules" block in the extension header is the spec;
 * the tests below lock each rule (split, merge, paste above/below,
 * cross-doc paste, undo/redo, change-tracker isolation).
 */

import { describe, test, expect } from "bun:test";
import { history, redo, undo } from "prosemirror-history";
import { Schema, Slice, Fragment } from "prosemirror-model";
import { EditorState } from "prosemirror-state";
import type { Command } from "prosemirror-state";

import {
  getChangedParagraphIds,
  ParagraphChangeTrackerExtension,
} from "./ParagraphChangeTrackerExtension";
import { ensureParaIdsInState, ParaIdAllocatorExtension } from "./ParaIdAllocatorExtension";

const schema = new Schema({
  nodes: {
    doc: { content: "block+" },
    paragraph: {
      group: "block",
      content: "inline*",
      attrs: {
        paraId: { default: null },
        textId: { default: null },
      },
      toDOM: () => ["p", 0],
    },
    text: { group: "inline" },
  },
});

const ext = ParaIdAllocatorExtension();
const runtime = ext.onSchemaReady({ schema });
const plugin = runtime.plugins?.[0];
if (!plugin) {
  throw new Error("Expected plugin from ParaIdAllocatorExtension");
}
const changeTrackerRuntime = ParagraphChangeTrackerExtension().onSchemaReady({
  schema,
});
const changeTrackerPlugin = changeTrackerRuntime.plugins?.[0];
if (!changeTrackerPlugin) {
  throw new Error("Expected plugin from ParagraphChangeTrackerExtension");
}

const para = (text: string, paraId: string | null = null) =>
  schema.node("paragraph", { paraId }, text.length > 0 ? [schema.text(text)] : []);

const createState = (...paras: ReturnType<typeof para>[]) =>
  EditorState.create({
    doc: schema.node("doc", null, paras),
    plugins: [plugin],
  });

const createTrackedState = (...paras: ReturnType<typeof para>[]) =>
  EditorState.create({
    doc: schema.node("doc", null, paras),
    plugins: [plugin, changeTrackerPlugin],
  });

const createHistoryState = (...paras: ReturnType<typeof para>[]) =>
  EditorState.create({
    doc: schema.node("doc", null, paras),
    plugins: [plugin, history()],
  });

const applyCommand = (state: EditorState, command: Command): EditorState => {
  let next = state;
  const applied = command(state, (tr) => {
    next = state.apply(tr);
  });
  expect(applied).toBe(true);
  return next;
};

const collectParaIds = (state: EditorState): (string | null)[] => {
  const out: (string | null)[] = [];
  state.doc.descendants((node) => {
    if (node.type.name === "paragraph") {
      const id = node.attrs["paraId"];
      out.push(typeof id === "string" ? id : null);
      return false;
    }
    return true;
  });
  return out;
};

describe("ParaIdAllocatorExtension", () => {
  test("allocates missing paraIds on the initial state before the first edit", () => {
    const initial = createState(para("Needs an id"), para("Also needs one"));
    expect(collectParaIds(initial)).toEqual([null, null]);

    const next = ensureParaIdsInState(initial);
    const ids = collectParaIds(next);

    expect(ids).toHaveLength(2);
    expect(ids[0]).toMatch(/^[0-9A-F]{8}$/u);
    expect(ids[1]).toMatch(/^[0-9A-F]{8}$/u);
    expect(ids[0]).not.toBe(ids[1]);
  });

  test("initial allocation does not seed paragraph change tracking", () => {
    const initial = createTrackedState(para("Needs an id"));

    const next = ensureParaIdsInState(initial);

    expect(collectParaIds(next)[0]).toMatch(/^[0-9A-F]{8}$/u);
    expect(getChangedParagraphIds(next).size).toBe(0);
  });

  test("initial allocation is idempotent when all paraIds are unique", () => {
    const initial = createState(para("First", "11111111"), para("Second", "22222222"));

    expect(ensureParaIdsInState(initial)).toBe(initial);
  });

  test("does not allocate on a selection-only transaction", () => {
    const initial = createState(para("Already has one", "ABCDEFGH"));
    const tr = initial.tr.setMeta("anything", true); // no doc change
    const next = initial.apply(tr);
    expect(collectParaIds(next)).toEqual(["ABCDEFGH"]);
  });

  test("allocates an 8-char hex id for a paragraph that lacks one", () => {
    const initial = createState(para("Needs an id"));
    // Insert a character to trigger a doc-change transaction and let
    // the appendTransaction hook fire.
    const next = initial.apply(initial.tr.insertText("!", 11));
    const ids = collectParaIds(next);
    expect(ids).toHaveLength(1);
    expect(ids[0]).toMatch(/^[0-9A-F]{8}$/u);
  });

  test("re-assigns a fresh id when a paragraph is inserted with a duplicate id", () => {
    const initial = createState(para("Original", "ABCDEFGH"));
    // Simulate paste: insert a paragraph node that carries the same
    // paraId as one already in the doc.
    const dupe = para("Pasted", "ABCDEFGH");
    const tr = initial.tr.replace(
      initial.doc.content.size,
      initial.doc.content.size,
      new Slice(Fragment.from(dupe), 0, 0),
    );
    const next = initial.apply(tr);
    const ids = collectParaIds(next);
    expect(ids).toHaveLength(2);
    expect(ids[0]).toBe("ABCDEFGH");
    expect(ids[1]).toMatch(/^[0-9A-F]{8}$/u);
    expect(ids[1]).not.toBe("ABCDEFGH");
  });

  test("preserves existing distinct ids untouched", () => {
    const initial = createState(para("First", "11111111"), para("Second", "22222222"));
    // Trigger any doc-changed transaction.
    const next = initial.apply(initial.tr.insertText("!", 6));
    expect(collectParaIds(next)).toEqual(["11111111", "22222222"]);
  });

  test("split: the first half keeps the id, the second half gets a fresh one", () => {
    const initial = createState(para("Hello", "5117AB1E"));
    // Split after "He": paragraph starts at 0, content at 1.
    const next = initial.apply(initial.tr.split(3));
    const ids = collectParaIds(next);
    expect(ids).toHaveLength(2);
    expect(ids[0]).toBe("5117AB1E");
    expect(ids[1]).toMatch(/^[0-9A-F]{8}$/u);
    expect(ids[1]).not.toBe("5117AB1E");
  });

  test("merge: the surviving paragraph keeps its id; the absorbed id dangles", () => {
    const initial = createState(para("First", "AAAA1111"), para("Second", "BBBB2222"));
    // Boundary between the two paragraphs: nodeSize of the first = 7.
    const next = initial.apply(initial.tr.join(7));
    expect(collectParaIds(next)).toEqual(["AAAA1111"]);
  });

  test("paste above the source: the source keeps its id, the pasted copy gets a fresh one", () => {
    const initial = createState(para("Original", "0R161NA1"));
    const dupe = para("Pasted copy", "0R161NA1");
    // Insert the duplicate BEFORE the source. Naive first-occurrence
    // dedupe would hand the pasted copy the source's id; the keeper
    // mapping must keep it on the original.
    const next = initial.apply(initial.tr.replace(0, 0, new Slice(Fragment.from(dupe), 0, 0)));
    const ids = collectParaIds(next);
    expect(ids).toHaveLength(2);
    // ids[0] is the pasted copy, ids[1] the original.
    expect(ids[1]).toBe("0R161NA1");
    expect(ids[0]).toMatch(/^[0-9A-F]{8}$/u);
    expect(ids[0]).not.toBe("0R161NA1");
    const texts: string[] = [];
    next.doc.descendants((node) => {
      if (node.type.name === "paragraph") {
        texts.push(node.textContent);
        return false;
      }
      return true;
    });
    expect(texts).toEqual(["Pasted copy", "Original"]);
  });

  test("paste of an id unknown to this doc is kept (cut-then-paste keeps anchors)", () => {
    const initial = createState(para("Existing", "EX157175"));
    const moved = para("Moved from elsewhere", "C0FFEE01");
    const next = initial.apply(
      initial.tr.replace(
        initial.doc.content.size,
        initial.doc.content.size,
        new Slice(Fragment.from(moved), 0, 0),
      ),
    );
    expect(collectParaIds(next)).toEqual(["EX157175", "C0FFEE01"]);
  });

  test("undo/redo: redoing a split mints a fresh id again (stable across saves, not redo)", () => {
    const initial = createHistoryState(para("Hello", "5117AB1E"));
    const afterSplit = initial.apply(initial.tr.split(3));
    const idsAfterSplit = collectParaIds(afterSplit);
    expect(idsAfterSplit[0]).toBe("5117AB1E");
    const mintedBeforeUndo = idsAfterSplit[1];
    expect(mintedBeforeUndo).toMatch(/^[0-9A-F]{8}$/u);

    const afterUndo = applyCommand(afterSplit, undo);
    expect(collectParaIds(afterUndo)).toEqual(["5117AB1E"]);

    const afterRedo = applyCommand(afterUndo, redo);
    const idsAfterRedo = collectParaIds(afterRedo);
    expect(idsAfterRedo).toHaveLength(2);
    expect(idsAfterRedo[0]).toBe("5117AB1E");
    // The redo-minted id is valid and unique, but NOT guaranteed to
    // equal the pre-undo id: the allocation transaction is outside
    // history (addToHistory: false), so redo re-creates the paragraph
    // and a fresh id is minted. Accepted behavior — an id is stable
    // across saves, not across redo-recreation.
    expect(idsAfterRedo[1]).toMatch(/^[0-9A-F]{8}$/u);
    expect(idsAfterRedo[1]).not.toBe("5117AB1E");
  });

  test("allocation via appendTransaction does not mark untouched paragraphs as changed", () => {
    const initial = createTrackedState(
      para("Never touched by the user"),
      para("Edited", "ED17ED01"),
    );
    // The user edits only the second paragraph; the allocator backfills
    // the first paragraph's missing id in the appended transaction.
    const editPos = initial.doc.content.size - 2;
    const next = initial.apply(initial.tr.insertText("!", editPos));

    const ids = collectParaIds(next);
    expect(ids[0]).toMatch(/^[0-9A-F]{8}$/u);
    expect(ids[1]).toBe("ED17ED01");

    const changed = getChangedParagraphIds(next);
    expect(changed.has("ED17ED01")).toBe(true);
    // SAFETY: asserted 8-hex above
    expect(changed.has(ids[0]!)).toBe(false);
  });
});
