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

import { describe, expect, spyOn, test } from "bun:test";
import { history, redo, undo } from "prosemirror-history";
import { Schema, Slice, Fragment } from "prosemirror-model";
import { EditorState } from "prosemirror-state";
import { TextSelection } from "prosemirror-state";
import type { Command } from "prosemirror-state";
import { ySyncPluginKey } from "y-prosemirror";

import { PROSE_PARAGRAPH_SOURCE_CONTRACT_ATTR } from "../../../docx/paragraphPropertySource";
import {
  ParagraphPropertySourceContract,
  ParagraphPropertySourceValidationError,
  paragraphPropertySourceTokenWire,
} from "../../../docx/paragraphPropertySourceIdentity";
import {
  joinParagraphsWithProperties,
  replaceSelectionThenSplitParagraphWithProperties,
  splitParagraphWithProperties,
} from "../../paragraphPropertyMutation";
import { EMPTY_EDITOR_PARAGRAPH_PROPERTY_STATE } from "../../paragraphPropertyState";
import {
  getChangedParagraphIds,
  ParagraphChangeTrackerExtension,
} from "./ParagraphChangeTrackerExtension";
import {
  ensureParaIdsInDoc,
  ensureParaIdsInState,
  ParaIdAllocatorExtension,
} from "./ParaIdAllocatorExtension";

const schema = new Schema({
  nodes: {
    doc: {
      content: "block+",
      attrs: { [PROSE_PARAGRAPH_SOURCE_CONTRACT_ATTR]: { default: null } },
    },
    blockquote: { group: "block", content: "block+" },
    paragraph: {
      group: "block",
      content: "inline*",
      attrs: {
        paraId: { default: null },
        textId: { default: null },
        _paragraphPropertyState: { default: EMPTY_EDITOR_PARAGRAPH_PROPERTY_STATE },
      },
      toDOM: () => ["p", 0],
    },
    text: { group: "inline" },
  },
  marks: { strong: {} },
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

const SOURCE_DIGEST = "a".repeat(64);
const SOURCE_CONTRACT = ParagraphPropertySourceContract.fromSourceCensus(SOURCE_DIGEST, [
  { paragraphCount: 10, story: { type: "document" } },
]).serialized;
const sourceToken = (ordinal: number) =>
  paragraphPropertySourceTokenWire({ type: "document" }, ordinal);
const sourcedPara = (text: string, paraId: string, ordinal: number) =>
  schema.node(
    "paragraph",
    {
      paraId,
      _paragraphPropertyState: {
        ...EMPTY_EDITOR_PARAGRAPH_PROPERTY_STATE,
        type: "imported",
        token: sourceToken(ordinal),
      },
    },
    text.length > 0 ? [schema.text(text)] : [],
  );

const createState = (...paras: ReturnType<typeof para>[]) =>
  EditorState.create({
    doc: schema.node("doc", null, paras),
    plugins: [plugin],
  });

const createSourcedState = (...paras: ReturnType<typeof sourcedPara>[]) =>
  EditorState.create({
    doc: schema.node("doc", { [PROSE_PARAGRAPH_SOURCE_CONTRACT_ATTR]: SOURCE_CONTRACT }, paras),
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

const createSourcedHistoryState = (...paras: ReturnType<typeof sourcedPara>[]) =>
  EditorState.create({
    doc: schema.node("doc", { [PROSE_PARAGRAPH_SOURCE_CONTRACT_ATTR]: SOURCE_CONTRACT }, paras),
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

const collectSourceTokens = (state: EditorState): (string | null)[] => {
  const out: (string | null)[] = [];
  state.doc.descendants((node) => {
    if (node.type.name === "paragraph") {
      const state = node.attrs["_paragraphPropertyState"];
      out.push(
        typeof state === "object" && state !== null && state.type === "imported"
          ? state.token
          : null,
      );
      return false;
    }
    return true;
  });
  return out;
};

describe("ParaIdAllocatorExtension", () => {
  test("allocates initial IDs by rebuilding only changed document branches", () => {
    const first = para("First", "11111111");
    const nested = schema.node("blockquote", null, [
      para("Duplicate", "11111111"),
      para("Missing"),
    ]);
    const original = schema.node("doc", null, [first, nested]);

    const normalized = ensureParaIdsInDoc(original);

    expect(normalized).not.toBe(original);
    expect(normalized.child(0)).toBe(first);
    expect(normalized.child(1)).not.toBe(nested);
    expect(normalized.child(1).child(0).attrs["paraId"]).toMatch(/^[0-9A-F]{8}$/u);
    expect(normalized.child(1).child(1).attrs["paraId"]).toMatch(/^[0-9A-F]{8}$/u);
    expect(normalized.child(1).child(0).attrs["paraId"]).not.toBe(
      normalized.child(1).child(1).attrs["paraId"],
    );
    expect(ensureParaIdsInDoc(normalized)).toBe(normalized);
  });

  test("reserves authored IDs before allocating missing initial IDs", () => {
    const random = spyOn(Math, "random");
    let calls = 0;
    try {
      random.mockImplementation(() => {
        calls += 1;
        return calls === 1 ? 0 : 0.5;
      });
      const original = schema.node("doc", null, [
        para("Missing"),
        para("Authored later", "00000001"),
      ]);

      const normalized = ensureParaIdsInDoc(original);

      expect(normalized.child(0).attrs["paraId"]).toBe("40000000");
      expect(normalized.child(1).attrs["paraId"]).toBe("00000001");
    } finally {
      random.mockRestore();
    }
  });

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

  test("replaces the reserved zero id carried by a pasted paragraph", () => {
    const initial = createState(para("Original", "ABCDEFGH"));
    const pasted = para("Pasted", "00000000");
    const next = initial.apply(
      initial.tr.replace(
        initial.doc.content.size,
        initial.doc.content.size,
        new Slice(Fragment.from(pasted), 0, 0),
      ),
    );

    const ids = collectParaIds(next);
    expect(ids[0]).toBe("ABCDEFGH");
    expect(ids[1]).toMatch(/^[0-9A-F]{8}$/u);
    expect(ids[1]).not.toBe("00000000");
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

  test("undo/redo: redoing a split restores the allocated id", () => {
    const random = spyOn(Math, "random");
    try {
      random.mockReturnValue(0);
      const initial = createHistoryState(para("Hello", "5117AB1E"));
      const afterSplit = initial.apply(initial.tr.split(3));
      const idsAfterSplit = collectParaIds(afterSplit);
      expect(idsAfterSplit[0]).toBe("5117AB1E");
      const mintedBeforeUndo = idsAfterSplit[1];
      expect(mintedBeforeUndo).toBe("00000001");

      const afterUndo = applyCommand(afterSplit, undo);
      expect(collectParaIds(afterUndo)).toEqual(["5117AB1E"]);

      const afterRedo = applyCommand(afterUndo, redo);
      const idsAfterRedo = collectParaIds(afterRedo);
      expect(idsAfterRedo).toHaveLength(2);
      expect(idsAfterRedo[0]).toBe("5117AB1E");
      expect(idsAfterRedo[1]).toBe(mintedBeforeUndo);
      expect(random).toHaveBeenCalledTimes(1);
    } finally {
      random.mockRestore();
    }
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

  test("native text and mark mutations retain the seeded source token", () => {
    const token = sourceToken(0);
    const initial = createSourcedState(sourcedPara("Hello", "AAAA1111", 0));
    const inserted = initial.apply(initial.tr.insertText("!", 3));
    const marked = inserted.apply(
      inserted.tr.addMark(1, inserted.doc.content.size - 1, schema.mark("strong")),
    );
    const unmarked = marked.apply(
      marked.tr.removeMark(1, marked.doc.content.size - 1, schema.marks.strong),
    );

    expect(collectSourceTokens(inserted)).toEqual([token]);
    expect(collectSourceTokens(marked)).toEqual([token]);
    expect(collectSourceTokens(unmarked)).toEqual([token]);
  });

  test.each(["", "Hello"])(
    "an attribute-only paragraph recreation cannot omit imported state for %j",
    (text) => {
      const initial = createSourcedState(sourcedPara(text, "AAAA1111", 0));
      expect(() =>
        initial.apply(
          initial.tr.setNodeMarkup(0, undefined, {
            paraId: "AAAA1111",
            textId: "BBBB2222",
          }),
        ),
      ).toThrow(ParagraphPropertySourceValidationError);
    },
  );

  test("a nested paragraph recreation cannot omit imported state", () => {
    const initial = EditorState.create({
      doc: schema.node("doc", { [PROSE_PARAGRAPH_SOURCE_CONTRACT_ATTR]: SOURCE_CONTRACT }, [
        schema.node("blockquote", null, [sourcedPara("Nested", "AAAA1111", 0)]),
      ]),
      plugins: [plugin],
    });
    expect(() =>
      initial.apply(
        initial.tr.setNodeMarkup(1, undefined, {
          paraId: "AAAA1111",
          textId: "BBBB2222",
        }),
      ),
    ).toThrow(ParagraphPropertySourceValidationError);
  });

  test.each([0, 2, 5])("a raw split at offset %i cannot duplicate imported ownership", (offset) => {
    const initial = createSourcedState(sourcedPara("Hello", "AAAA1111", 0));
    expect(() => initial.apply(initial.tr.split(offset + 1))).toThrow(
      ParagraphPropertySourceValidationError,
    );
  });

  test("the ownership proof step survives split undo and redo", () => {
    const initial = createSourcedHistoryState(sourcedPara("Hello", "AAAA1111", 0));
    const transaction = initial.tr;
    splitParagraphWithProperties({
      transaction,
      pos: 3,
      transition: { type: "split-left-created-right-retains" },
    });
    const split = initial.apply(transaction);
    const undone = applyCommand(split, undo);
    const redone = applyCommand(undone, redo);

    expect(collectSourceTokens(split)).toEqual([null, sourceToken(0)]);
    expect(collectSourceTokens(undone)).toEqual([sourceToken(0)]);
    expect(collectSourceTokens(redone)).toEqual([null, sourceToken(0)]);
  });

  test("replace-selection and split records ownership before either structural step", () => {
    const initial = createSourcedHistoryState(
      sourcedPara("First", "AAAA1111", 0),
      sourcedPara("Second", "BBBB2222", 1),
    );
    const transaction = initial.tr.setSelection(TextSelection.create(initial.doc, 3, 10));
    replaceSelectionThenSplitParagraphWithProperties({
      transaction,
      transition: { type: "split-left-created-right-retains" },
    });
    const changed = initial.apply(transaction);
    const undone = applyCommand(changed, undo);
    const redone = applyCommand(undone, redo);

    expect(collectSourceTokens(changed)).toEqual([null, sourceToken(0)]);
    expect(collectSourceTokens(undone)).toEqual([sourceToken(0), sourceToken(1)]);
    expect(collectSourceTokens(redone)).toEqual([null, sourceToken(0)]);
  });

  test.each([0, 2, 5])(
    "the canonical split at offset %i carries an issuer proof and retains the right owner",
    (offset) => {
      const initial = createSourcedState(sourcedPara("Hello", "AAAA1111", 0));
      const transaction = initial.tr;
      splitParagraphWithProperties({
        transaction,
        pos: offset + 1,
        transition: { type: "split-left-created-right-retains" },
      });
      const next = initial.apply(transaction);

      expect(collectSourceTokens(next)).toEqual([null, sourceToken(0)]);
    },
  );

  test("pasting a duplicate above the source cannot steal its token", () => {
    const original = sourcedPara("Original", "AAAA1111", 0);
    const initial = createSourcedState(original);
    const copy = original.type.create(
      { ...original.attrs, paraId: "BBBB2222" },
      schema.text("Copy"),
    );
    expect(() =>
      initial.apply(initial.tr.replace(0, 0, new Slice(Fragment.from(copy), 0, 0))),
    ).toThrow(ParagraphPropertySourceValidationError);
  });

  test("a contract-owned token cannot appear without an ownership transition", () => {
    const initial = createSourcedState(sourcedPara("Original", "AAAA1111", 0));
    const foreign = schema.node(
      "paragraph",
      {
        paraId: "BBBB2222",
        _paragraphPropertyState: {
          ...EMPTY_EDITOR_PARAGRAPH_PROPERTY_STATE,
          type: "imported",
          token: sourceToken(9),
        },
      },
      schema.text("Foreign"),
    );
    expect(() =>
      initial.apply(
        initial.tr.replace(
          initial.doc.content.size,
          initial.doc.content.size,
          new Slice(Fragment.from(foreign), 0, 0),
        ),
      ),
    ).toThrow(ParagraphPropertySourceValidationError);
  });

  test("a valid-looking token outside the contract census is rejected", () => {
    const initial = createSourcedState(sourcedPara("Original", "AAAA1111", 0));
    const forged = schema.node(
      "paragraph",
      {
        paraId: "BBBB2222",
        _paragraphPropertyState: {
          ...EMPTY_EDITOR_PARAGRAPH_PROPERTY_STATE,
          type: "imported",
          token: sourceToken(10),
        },
      },
      schema.text("Forged"),
    );

    expect(() =>
      initial.apply(
        initial.tr.replace(
          initial.doc.content.size,
          initial.doc.content.size,
          new Slice(Fragment.from(forged), 0, 0),
        ),
      ),
    ).toThrow(ParagraphPropertySourceValidationError);
  });

  test("a token from another contract story cannot enter the editable body", () => {
    const contract = ParagraphPropertySourceContract.fromSourceCensus(SOURCE_DIGEST, [
      { paragraphCount: 1, story: { type: "document" } },
      { paragraphCount: 1, story: { relationshipId: "rId1", type: "header" } },
    ]);
    const headerToken = paragraphPropertySourceTokenWire(
      { relationshipId: "rId1", type: "header" },
      0,
    );
    const headerOwned = schema.node(
      "paragraph",
      {
        paraId: "AAAA1111",
        _paragraphPropertyState: {
          ...EMPTY_EDITOR_PARAGRAPH_PROPERTY_STATE,
          type: "imported",
          token: headerToken,
        },
      },
      schema.text("Header"),
    );

    expect(() =>
      EditorState.create({
        doc: schema.node("doc", { [PROSE_PARAGRAPH_SOURCE_CONTRACT_ATTR]: contract.serialized }, [
          headerOwned,
        ]),
        plugins: [plugin],
      }),
    ).toThrow(ParagraphPropertySourceValidationError);
  });

  test("a raw paragraph deletion cannot discard imported ownership", () => {
    const original = sourcedPara("Original", "AAAA1111", 0);
    const sibling = sourcedPara("Sibling", "BBBB2222", 1);
    const initial = createSourcedState(original, sibling);
    expect(() => initial.apply(initial.tr.delete(0, original.nodeSize))).toThrow(
      ParagraphPropertySourceValidationError,
    );
  });

  test("a remote deletion is accepted only when the mapping deletes the owner interior", () => {
    const original = sourcedPara("Original", "AAAA1111", 0);
    const sibling = sourcedPara("Sibling", "BBBB2222", 1);
    const initial = createSourcedState(original, sibling);
    const transaction = initial.tr
      .delete(0, original.nodeSize)
      .setMeta(ySyncPluginKey, { type: "synthetic-remote-boundary" });
    const next = initial.apply(transaction);

    expect(collectSourceTokens(next)).toEqual([sourceToken(1)]);
  });

  test("a remote undo can restore an exact previously observed deleted owner", () => {
    const original = sourcedPara("Original", "AAAA1111", 0);
    const sibling = sourcedPara("Sibling", "BBBB2222", 1);
    const initial = createSourcedState(original, sibling);
    const deleted = initial.apply(
      initial.tr
        .delete(0, original.nodeSize)
        .setMeta(ySyncPluginKey, { type: "synthetic-remote-boundary" }),
    );
    const restored = deleted.apply(
      deleted.tr
        .replace(0, 0, new Slice(Fragment.from(original), 0, 0))
        .setMeta(ySyncPluginKey, { type: "synthetic-remote-boundary" }),
    );

    expect(collectSourceTokens(restored)).toEqual([sourceToken(0), sourceToken(1)]);
  });

  test("a remote insertion cannot mint an unobserved contract-owned owner", () => {
    const initial = createSourcedState(sourcedPara("Original", "AAAA1111", 0));
    const unseen = sourcedPara("Unseen", "BBBB2222", 9);
    const transaction = initial.tr
      .replace(
        initial.doc.content.size,
        initial.doc.content.size,
        new Slice(Fragment.from(unseen), 0, 0),
      )
      .setMeta(ySyncPluginKey, { type: "synthetic-remote-boundary" });

    expect(() => initial.apply(transaction)).toThrow(ParagraphPropertySourceValidationError);
  });

  test("a remote split cannot reuse an imported owner without a structural proof", () => {
    const initial = createSourcedState(sourcedPara("Original", "AAAA1111", 0));
    const transaction = initial.tr
      .split(4)
      .setMeta(ySyncPluginKey, { type: "synthetic-remote-boundary" });

    expect(() => initial.apply(transaction)).toThrow(ParagraphPropertySourceValidationError);
  });

  test("a remote split is reified from an exact content partition and owner transition", () => {
    const initial = createSourcedState(sourcedPara("Original", "AAAA1111", 0));
    const transaction = initial.tr.split(4);
    const left = transaction.doc.child(0);
    transaction.setNodeMarkup(0, undefined, {
      ...left.attrs,
      _paragraphPropertyState: EMPTY_EDITOR_PARAGRAPH_PROPERTY_STATE,
    });
    transaction.setMeta(ySyncPluginKey, { type: "synthetic-remote-boundary" });
    const next = initial.apply(transaction);

    expect(collectSourceTokens(next)).toEqual([null, sourceToken(0)]);
  });

  test("a remote owner relocation without an exact split or join topology is rejected", () => {
    const initial = createSourcedState(sourcedPara("Original", "AAAA1111", 0));
    const original = initial.doc.child(0);
    const detached = original.type.create(
      {
        ...original.attrs,
        _paragraphPropertyState: EMPTY_EDITOR_PARAGRAPH_PROPERTY_STATE,
      },
      original.content,
      original.marks,
    );
    const forged = sourcedPara("Different", "BBBB2222", 0);
    const transaction = initial.tr
      .replaceWith(0, original.nodeSize, [detached, forged])
      .setMeta(ySyncPluginKey, { type: "synthetic-remote-boundary" });

    expect(() => initial.apply(transaction)).toThrow(ParagraphPropertySourceValidationError);
  });

  test("a remote whole-paragraph join is reified from adjacent exact content", () => {
    const left = sourcedPara("Left", "AAAA1111", 0);
    const initial = createSourcedState(left, sourcedPara("Right", "BBBB2222", 1));
    const transaction = initial.tr
      .join(left.nodeSize)
      .setMeta(ySyncPluginKey, { type: "synthetic-remote-boundary" });
    const next = initial.apply(transaction);

    expect(collectSourceTokens(next)).toEqual([sourceToken(0)]);
  });

  test("the explicit revision join can select the right paragraph owner", () => {
    const initial = createSourcedState(
      sourcedPara("", "AAAA1111", 0),
      sourcedPara("Right", "BBBB2222", 1),
    );
    const transaction = initial.tr;
    joinParagraphsWithProperties({
      joinPos: transaction.doc.child(0).nodeSize,
      transaction,
      transition: { type: "join-right-paragraph-mark-retains" },
    });
    const next = initial.apply(transaction);

    expect(collectSourceTokens(next)).toEqual([sourceToken(1)]);
  });
});
