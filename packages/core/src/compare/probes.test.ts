/**
 * Twelve labelled single mutations, one probe each.
 *
 * The property tests generate scripts and pin what must hold for all of them.
 * These pin the opposite thing: for one named edit a reviewer would recognize,
 * exactly what the change list should say. A comparison that round-trips
 * correctly can still describe a split paragraph as two rewrites, or a bolded
 * phrase as a deletion and a reinsertion of the same words, and a reader
 * cannot tell those apart from the real edit. The property suite cannot see
 * that; a named probe can.
 *
 * Each probe applies its mutation directly to a base (so the difference is
 * known by construction), compares, and asserts the round trip plus the change
 * kinds. Where the engine does not yet describe the edit the way a reviewer
 * would, the probe records what it does describe and names the gap, rather
 * than being deleted or weakened into a tautology.
 */

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";
import { readFileSync } from "node:fs";
import path from "node:path";

import { FolioDocxReviewer } from "../ai-edits/headless";
import type { FolioAIBlock } from "../ai-edits/types";
import { buildBodySequenceDocx } from "./__fixtures__/body-sequence";
import { buildNestedTableDocx } from "./__fixtures__/nested-table";
import {
  buildNumberedListDocx,
  NUMBERED_LIST_ITEMS,
  withItemDemoted,
  withItemUnnumbered,
} from "./__fixtures__/numbered-list";
import { compareDocx } from "./compare";
import { applyEditScript, type EditScript } from "./scenario";
import type { CompareChange } from "./types";

const FIXTURES_DIR = path.join(import.meta.dir, "../docx/__tests__/__fixtures__/corpus");

const readFixture = (filename: string): ArrayBuffer => {
  const bytes = readFileSync(path.join(FIXTURES_DIR, filename));
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
};

const OPTIONS = { author: "compare", timestamp: "2024-03-01T00:00:00.000Z" } as const;

const PROSE_BASE = readFixture("upstream-styled-content.docx");
const TABLE_BASE = readFixture("upstream-with-tables.docx");
/** Authored here: no corpus fixture carries numbering. */
const LIST_BASE = await buildNumberedListDocx();

const documentPartOf = async (buffer: ArrayBuffer): Promise<string> =>
  (await (await JSZip.loadAsync(buffer)).file("word/document.xml")?.async("string")) ?? "";

const blocksOf = async (buffer: ArrayBuffer): Promise<FolioAIBlock[]> =>
  (await FolioDocxReviewer.fromBuffer(buffer)).getContent();

type BlockProjection = { text: string; table: FolioAIBlock["table"] | null };

const projectView = async (
  buffer: ArrayBuffer,
  view: "original" | "final",
): Promise<BlockProjection[]> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(buffer);
  const story = reviewer.readReviewedStory({ view });
  return (story?.snapshot.blocks ?? []).map((block) => ({
    text: block.text,
    table: block.table ?? null,
  }));
};

type ProbeOutcome = {
  changes: readonly CompareChange[];
  kinds: readonly string[];
};

/**
 * Apply one mutation, compare, and assert the round trip. Every probe shares
 * this much; what differs is what each one then says about the change list.
 */
const probe = async (base: ArrayBuffer, script: EditScript): Promise<ProbeOutcome> => {
  const scripted = await applyEditScript(base, script);
  if (scripted.isErr()) {
    throw scripted.error;
  }
  expect(scripted.value.unresolved).toEqual([]);
  const target = scripted.value.buffer;

  const result = await compareDocx(base, target, OPTIONS);
  if (result.isErr()) {
    throw result.error;
  }

  expect(await projectView(result.value.buffer, "final")).toEqual(
    await projectView(target, "final"),
  );
  expect(await projectView(result.value.buffer, "original")).toEqual(
    await projectView(base, "final"),
  );

  return { changes: result.value.changes, kinds: result.value.changes.map(({ kind }) => kind) };
};

/** A block with enough words to split, edit inside, or bold part of. */
const wordyBlockIndex = (blocks: readonly FolioAIBlock[], minimumWords: number): number => {
  const index = blocks.findIndex(
    (block) => !block.table && block.text.split(" ").length >= minimumWords,
  );
  if (index === -1) {
    throw new Error("The fixture has no paragraph long enough for this probe.");
  }
  return index;
};

const firstTableBlockIndex = (blocks: readonly FolioAIBlock[]): number => {
  const index = blocks.findIndex((block) => block.table !== undefined);
  if (index === -1) {
    throw new Error("The fixture has no table.");
  }
  return index;
};

const PROSE_BLOCKS = await blocksOf(PROSE_BASE);
const TABLE_BLOCKS = await blocksOf(TABLE_BASE);
const LIST_BLOCKS = await blocksOf(LIST_BASE);

describe("single-mutation probes", () => {
  test("insert_sentence: one added sentence is one replace", async () => {
    const blockIndex = wordyBlockIndex(PROSE_BLOCKS, 4);
    const word = PROSE_BLOCKS[blockIndex]?.text.split(" ").at(0) ?? "";
    const { kinds } = await probe(PROSE_BASE, [
      { type: "replaceWords", blockIndex, find: word, replace: `${word} A new sentence applies.` },
    ]);
    expect(kinds).toEqual(["replace"]);
  });

  test("delete_sentence: one removed phrase is one replace", async () => {
    const blockIndex = wordyBlockIndex(PROSE_BLOCKS, 6);
    const phrase = PROSE_BLOCKS[blockIndex]?.text.split(" ").slice(1, 3).join(" ") ?? "";
    const { kinds } = await probe(PROSE_BASE, [
      { type: "replaceWords", blockIndex, find: `${phrase} `, replace: "" },
    ]);
    expect(kinds).toEqual(["replace"]);
  });

  test("insert_paragraph: a new paragraph is one insert", async () => {
    const { kinds } = await probe(PROSE_BASE, [
      {
        type: "insertParagraphAfter",
        blockIndex: wordyBlockIndex(PROSE_BLOCKS, 3),
        text: "An entirely new clause governs the schedule.",
      },
    ]);
    expect(kinds).toEqual(["insert"]);
  });

  test("delete_paragraph: a removed paragraph is one delete", async () => {
    const { kinds } = await probe(PROSE_BASE, [
      { type: "deleteParagraph", blockIndex: wordyBlockIndex(PROSE_BLOCKS, 3) },
    ]);
    expect(kinds).toEqual(["delete"]);
  });

  /**
   * A split moves a paragraph mark and no words, and is reported as exactly
   * that: one inserted mark. Reporting it as a rewrite of the head plus an
   * insertion of the tail round-trips and lies — the tail was not written
   * today.
   */
  test("split_paragraph: one inserted paragraph mark", async () => {
    const blockIndex = wordyBlockIndex(PROSE_BLOCKS, 6);
    const { changes, kinds } = await probe(PROSE_BASE, [
      { type: "splitParagraph", blockIndex, wordIndex: 3 },
    ]);
    expect(kinds).toEqual(["split"]);
    const [change] = changes;
    expect(change?.kind === "split" && change.targetBlockIds).toHaveLength(2);
    expect(change?.kind === "split" && change.text).toBe(PROSE_BLOCKS[blockIndex]?.text);
  });

  /** The mirror of the split: one deleted paragraph mark. */
  test("merge_paragraphs: one deleted paragraph mark", async () => {
    const blockIndex = wordyBlockIndex(PROSE_BLOCKS, 3);
    const { changes, kinds } = await probe(PROSE_BASE, [{ type: "mergeParagraphs", blockIndex }]);
    expect(kinds).toEqual(["merge"]);
    const [change] = changes;
    expect(change?.kind === "merge" && change.baseBlockIds).toHaveLength(2);
  });

  test("move_clause: a relocated paragraph is reported as a move", async () => {
    const blockIndex = wordyBlockIndex(PROSE_BLOCKS, 4);
    const beforeBlockIndex = PROSE_BLOCKS.length - 1;
    const { kinds } = await probe(PROSE_BASE, [
      { type: "moveParagraph", blockIndex, beforeBlockIndex },
    ]);
    // The alignment may absorb a relocation it can still walk forward past;
    // what it must not do is report it as unrelated churn.
    expect(kinds.every((kind) => kind === "move")).toBe(true);
  });

  test("move_clause: the package carries a linked w:moveFrom / w:moveTo pair", async () => {
    // A move reported only in the change list is a move the document does not
    // know about: every OOXML consumer sees an unrelated deletion and
    // insertion, and a reviewer reading the redline in Word cannot tell the
    // text was relocated rather than rewritten.
    const blockIndex = wordyBlockIndex(PROSE_BLOCKS, 4);
    const scripted = await applyEditScript(PROSE_BASE, [
      { type: "moveParagraph", blockIndex, beforeBlockIndex: PROSE_BLOCKS.length - 1 },
    ]);
    if (scripted.isErr()) {
      throw scripted.error;
    }
    const result = await compareDocx(PROSE_BASE, scripted.value.buffer, OPTIONS);
    if (result.isErr()) {
      throw result.error;
    }
    expect(result.value.changes.map(({ kind }) => kind)).toEqual(["move"]);

    const documentXml = await documentPartOf(result.value.buffer);
    expect(documentXml).toContain("<w:moveFrom ");
    expect(documentXml).toContain("<w:moveTo ");
    // The relocated text is not also written as a plain insertion or deletion.
    expect(documentXml).not.toContain("<w:ins ");
    expect(documentXml).not.toContain("<w:del ");

    // A move pair resolves like any other revision.
    expect(await projectView(result.value.buffer, "final")).toEqual(
      await projectView(scripted.value.buffer, "final"),
    );
    expect(await projectView(result.value.buffer, "original")).toEqual(
      await projectView(PROSE_BASE, "final"),
    );
  });

  test("add_list_item: an added item is one insert", async () => {
    const { kinds } = await probe(LIST_BASE, [
      { type: "insertParagraphAfter", blockIndex: 1, text: "Delivery may be made in instalments." },
    ]);
    expect(kinds).toEqual(["insert"]);
  });

  test("delete_table_row: a removed row is one table-row-delete", async () => {
    const { kinds } = await probe(TABLE_BASE, [
      { type: "deleteTableRow", blockIndex: firstTableBlockIndex(TABLE_BLOCKS) },
    ]);
    expect(kinds).toEqual(["table-row-delete"]);
  });

  /**
   * The story's last block sits in a table inside a table, so the appended
   * paragraph's only anchor is two cells deep. A block insertion escapes the
   * table it is anchored in; escaping one level leaves it in the outer cell,
   * where the paragraph is not a document-level peer and the round trip fails.
   */
  test("append_after_nested_table: the new paragraph lands at body level", async () => {
    const base = await buildNestedTableDocx();
    const target = await buildNestedTableDocx({
      trailingParagraph: "Signed by the parties on the date first written above.",
    });

    const result = await compareDocx(base, target, OPTIONS);
    if (result.isErr()) {
      throw result.error;
    }
    expect(result.value.changes.map(({ kind }) => kind)).toEqual(["insert"]);
    expect(await projectView(result.value.buffer, "final")).toEqual(
      await projectView(target, "final"),
    );
    expect(await projectView(result.value.buffer, "original")).toEqual(
      await projectView(base, "final"),
    );
  });

  test("granularity: character-level marks less of the word than word-level does", async () => {
    const blockIndex = wordyBlockIndex(PROSE_BLOCKS, 4);
    const word = PROSE_BLOCKS[blockIndex]?.text.split(" ").at(0) ?? "";
    const scripted = await applyEditScript(PROSE_BASE, [
      { type: "replaceWords", blockIndex, find: word, replace: `${word}ing` },
    ]);
    if (scripted.isErr()) {
      throw scripted.error;
    }
    const target = scripted.value.buffer;

    const wordLevel = await compareDocx(PROSE_BASE, target, OPTIONS);
    const characterLevel = await compareDocx(PROSE_BASE, target, {
      ...OPTIONS,
      granularity: "character",
    });
    if (wordLevel.isErr() || characterLevel.isErr()) {
      throw wordLevel.isErr() ? wordLevel.error : characterLevel.error;
    }

    // The option reaches the package: the same edit is redlined two ways.
    expect(new Uint8Array(wordLevel.value.buffer)).not.toEqual(
      new Uint8Array(characterLevel.value.buffer),
    );
    // Both still accept back to the target and reject back to the base.
    for (const result of [wordLevel.value, characterLevel.value]) {
      expect(await projectView(result.buffer, "final")).toEqual(await projectView(target, "final"));
      expect(await projectView(result.buffer, "original")).toEqual(
        await projectView(PROSE_BASE, "final"),
      );
    }
  });

  test("format_only_bold: bolding a phrase is a format, never a rewrite", async () => {
    const blockIndex = wordyBlockIndex(PROSE_BLOCKS, 4);
    const { changes, kinds } = await probe(PROSE_BASE, [
      { type: "formatRange", blockIndex, startOffset: 0, endOffset: 4, formatting: { bold: true } },
    ]);
    expect(kinds).toEqual(["format"]);
    const [change] = changes;
    expect(change?.kind === "format" && change.ranges.length).toBe(1);
  });

  test("renumbering: an added list item does not report the items after it", async () => {
    // Inserting at the top renumbers every item below. Labels are rendered
    // from the numbering definitions rather than stored in the paragraphs, so
    // the comparison must report the insertion and nothing else; reporting
    // the renumbered items would bury the real edit.
    expect(LIST_BLOCKS.length).toBe(NUMBERED_LIST_ITEMS.length);
    const { kinds } = await probe(LIST_BASE, [
      {
        type: "insertParagraphAfter",
        blockIndex: 0,
        text: "The goods shall conform to the specification.",
      },
    ]);
    expect(kinds).toEqual(["insert"]);
  });

  test("renumbering_definition: a changed list format is reported, not silently dropped", async () => {
    // The words are identical; only `numbering.xml` differs. Labels are
    // rendered from the definitions, so nothing in any block's text moves and
    // a text-only comparison sees two identical documents.
    const roman = await buildNumberedListDocx(NUMBERED_LIST_ITEMS, { format: "lowerRoman" });
    const result = await compareDocx(LIST_BASE, roman, OPTIONS);
    if (result.isErr()) {
      throw result.error;
    }
    const kinds = result.value.changes.map(({ kind }) => kind);
    expect(new Set(kinds)).toEqual(new Set(["numbering"]));
    const [change] = result.value.changes;
    expect(change?.kind === "numbering" && change.before?.format).toBe("decimal");
    expect(change?.kind === "numbering" && change.after?.format).toBe("lowerRoman");
  });

  test("change_list_level: a demoted list item is one paragraph-format change", async () => {
    // Demoting an item changes `w:ilvl` and nothing a text diff can see. It
    // used to reach the comparison as no change at all, so the redline said
    // the two documents agreed; it is now a `w:pPrChange`, which is what Word
    // writes for the same edit.
    const demoted = await buildNumberedListDocx(withItemDemoted(NUMBERED_LIST_ITEMS, 3));
    const result = await compareDocx(LIST_BASE, demoted, OPTIONS);
    if (result.isErr()) {
      throw result.error;
    }
    expect(result.value.changes.map(({ kind }) => kind)).toEqual(["paragraph-format"]);
    const [change] = result.value.changes;
    expect(change?.kind === "paragraph-format" && change.properties).toEqual({ listLevel: 1 });

    expect(await projectView(result.value.buffer, "final")).toEqual(
      await projectView(demoted, "final"),
    );
    expect(await projectView(result.value.buffer, "original")).toEqual(
      await projectView(LIST_BASE, "final"),
    );
    expect(await documentPartOf(result.value.buffer)).toContain("<w:pPrChange ");
  });

  test("unnumber_list_item: a paragraph that stopped being a list item is reported", async () => {
    // The words are identical and only `w:numPr` is gone. Reading the target's
    // level alone made this invisible — an absent level is not a level that
    // differs — so the comparison reported nothing and its own self-check
    // then refused the pair, because the accepted result was still a list.
    const unnumbered = await buildNumberedListDocx(withItemUnnumbered(NUMBERED_LIST_ITEMS, 2));
    const result = await compareDocx(LIST_BASE, unnumbered, OPTIONS);
    if (result.isErr()) {
      throw result.error;
    }
    expect(result.value.changes.map(({ kind }) => kind)).toEqual(["paragraph-format"]);
    const [change] = result.value.changes;
    expect(change?.kind === "paragraph-format" && change.properties).toEqual({ listLevel: null });

    expect(await projectView(result.value.buffer, "final")).toEqual(
      await projectView(unnumbered, "final"),
    );
    expect(await projectView(result.value.buffer, "original")).toEqual(
      await projectView(LIST_BASE, "final"),
    );
    expect(await documentPartOf(result.value.buffer)).toContain("<w:pPrChange ");
  });

  test("append_paragraph_then_table: additions past the last block keep target order", async () => {
    // Both additions resolve to the same position — after the base's last
    // block — so their order in the document is their order in the operation
    // list and nothing else. Collecting the paragraph insertions and emitting
    // them after the loop put every one of them behind a table that the target
    // has after them.
    const base = await buildBodySequenceDocx([
      { kind: "paragraph", text: "The parties agree as set out below." },
      { kind: "paragraph", text: "This agreement is governed by the stated law." },
    ]);
    const target = await buildBodySequenceDocx([
      { kind: "paragraph", text: "The parties agree as set out below." },
      { kind: "paragraph", text: "This agreement is governed by the stated law." },
      { kind: "paragraph", text: "The schedule below records the agreed fees." },
      { kind: "table", rows: [["Service", "Fee"]] },
    ]);

    const result = await compareDocx(base, target, OPTIONS);
    if (result.isErr()) {
      throw result.error;
    }
    expect(await projectView(result.value.buffer, "final")).toEqual(
      await projectView(target, "final"),
    );
    expect(await projectView(result.value.buffer, "original")).toEqual(
      await projectView(base, "final"),
    );
  });
});
