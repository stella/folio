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
import type { Command } from "prosemirror-state";

import {
  PROSE_PARAGRAPH_SOURCE_CONTRACT_ATTR,
  PROSE_PARAGRAPH_SOURCE_TOKEN_ATTR,
  joinProseParagraphsWithRightPropertySource,
} from "../../../docx/paragraphPropertySource";
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
        [PROSE_PARAGRAPH_SOURCE_TOKEN_ATTR]: { default: null },
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
const SOURCE_CONTRACT = `folio-ppr-v1:${SOURCE_DIGEST}`;
const sourceToken = (ordinal: number, fingerprint = SOURCE_DIGEST.slice(0, 32)) =>
  `p1d:${fingerprint}:${ordinal.toString(36)}`;
const sourcedPara = (text: string, paraId: string, ordinal: number) =>
  schema.node(
    "paragraph",
    { paraId, [PROSE_PARAGRAPH_SOURCE_TOKEN_ATTR]: sourceToken(ordinal) },
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
      const token = node.attrs[PROSE_PARAGRAPH_SOURCE_TOKEN_ATTR];
      out.push(typeof token === "string" ? token : null);
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
    "an attribute-only paragraph recreation restores an omitted seeded source token for %j",
    (text) => {
      const initial = createSourcedState(sourcedPara(text, "AAAA1111", 0));
      const next = initial.apply(
        initial.tr.setNodeMarkup(0, undefined, {
          paraId: "AAAA1111",
          textId: "BBBB2222",
        }),
      );

      expect(collectSourceTokens(next)).toEqual([sourceToken(0)]);
    },
  );

  test("a nested paragraph recreation restores an omitted seeded source token", () => {
    const initial = EditorState.create({
      doc: schema.node("doc", { [PROSE_PARAGRAPH_SOURCE_CONTRACT_ATTR]: SOURCE_CONTRACT }, [
        schema.node("blockquote", null, [sourcedPara("Nested", "AAAA1111", 0)]),
      ]),
      plugins: [plugin],
    });
    const next = initial.apply(
      initial.tr.setNodeMarkup(1, undefined, {
        paraId: "AAAA1111",
        textId: "BBBB2222",
      }),
    );

    expect(collectSourceTokens(next)).toEqual([sourceToken(0)]);
  });

  test.each([0, 2, 5])(
    "split at offset %i retains the token only on the mapped original half",
    (offset) => {
      const initial = createSourcedState(sourcedPara("Hello", "AAAA1111", 0));
      const next = initial.apply(initial.tr.split(offset + 1));

      expect(collectSourceTokens(next)).toEqual([sourceToken(0), null]);
    },
  );

  test("pasting a duplicate above the source cannot steal its token", () => {
    const original = sourcedPara("Original", "AAAA1111", 0);
    const initial = createSourcedState(original);
    const copy = original.type.create(
      { ...original.attrs, paraId: "BBBB2222" },
      schema.text("Copy"),
    );
    const next = initial.apply(initial.tr.replace(0, 0, new Slice(Fragment.from(copy), 0, 0)));

    expect(collectSourceTokens(next)).toEqual([null, sourceToken(0)]);
  });

  test("a token from a different source fingerprint is cleared", () => {
    const initial = createSourcedState(sourcedPara("Original", "AAAA1111", 0));
    const foreign = schema.node(
      "paragraph",
      {
        paraId: "BBBB2222",
        [PROSE_PARAGRAPH_SOURCE_TOKEN_ATTR]: sourceToken(0, "b".repeat(32)),
      },
      schema.text("Foreign"),
    );
    const next = initial.apply(
      initial.tr.replace(
        initial.doc.content.size,
        initial.doc.content.size,
        new Slice(Fragment.from(foreign), 0, 0),
      ),
    );

    expect(collectSourceTokens(next)).toEqual([sourceToken(0), null]);
  });

  test("a uniquely seeded token survives delete then paste", () => {
    const original = sourcedPara("Original", "AAAA1111", 0);
    const sibling = sourcedPara("Sibling", "BBBB2222", 1);
    const initial = createSourcedState(original, sibling);
    const deleted = initial.apply(initial.tr.delete(0, original.nodeSize));
    const restored = deleted.apply(
      deleted.tr.replace(
        deleted.doc.content.size,
        deleted.doc.content.size,
        new Slice(Fragment.from(original), 0, 0),
      ),
    );

    expect(collectSourceTokens(deleted)).toEqual([sourceToken(1)]);
    expect(collectSourceTokens(restored)).toEqual([sourceToken(1), sourceToken(0)]);
  });

  test("the explicit revision join can select the right paragraph owner", () => {
    const initial = createSourcedState(
      sourcedPara("", "AAAA1111", 0),
      sourcedPara("Right", "BBBB2222", 1),
    );
    const transaction = initial.tr;
    joinProseParagraphsWithRightPropertySource({
      attrs: transaction.doc.child(1).attrs,
      pos: transaction.doc.child(0).nodeSize,
      transaction,
    });
    const next = initial.apply(transaction);

    expect(collectSourceTokens(next)).toEqual([sourceToken(1)]);
  });
});
