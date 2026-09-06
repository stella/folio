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
import { createDocx } from "../docx/rezip";
import type { Table, TableCell } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
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

type ColumnCell = { text: string; gridSpan?: number };

const buildColumnTableDocx = (rows: readonly (readonly ColumnCell[])[]): Promise<ArrayBuffer> => {
  const template = createEmptyDocument();
  const cell = ({ text, gridSpan }: ColumnCell): TableCell => ({
    type: "tableCell",
    ...(gridSpan !== undefined && { formatting: { gridSpan } }),
    content: [
      {
        type: "paragraph",
        content: [{ type: "run", content: [{ type: "text", text }] }],
      },
    ],
  });
  const table: Table = {
    type: "table",
    rows: rows.map((cells) => ({ type: "tableRow", cells: cells.map(cell) })),
  };
  return createDocx({
    ...template,
    package: {
      ...template.package,
      document: { ...template.package.document, content: [table] },
    },
  });
};

const documentPartOf = async (buffer: ArrayBuffer): Promise<string> =>
  (await (await JSZip.loadAsync(buffer)).file("word/document.xml")?.async("string")) ?? "";

/** Every `w:p` element of a document part, as raw XML. */
const paragraphsOf = (documentXml: string): string[] =>
  documentXml.match(/<w:p[ >][\s\S]*?<\/w:p>/gu) ?? [];

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

  test("insert_table_column: a grid-aligned column is inserted beside merged neighbours", async () => {
    const base = await buildColumnTableDocx([
      [{ text: "Account details", gridSpan: 2 }, { text: "Status" }],
      [{ text: "Fees" }, { text: "Annual" }, { text: "Open" }],
    ]);
    const target = await buildColumnTableDocx([
      [{ text: "Account details", gridSpan: 2 }, { text: "Currency" }, { text: "Status" }],
      [{ text: "Fees" }, { text: "Annual" }, { text: "EUR" }, { text: "Open" }],
    ]);

    const result = await compareDocx(base, target, OPTIONS);
    if (result.isErr()) {
      throw result.error;
    }
    expect(result.value.verification).toEqual({ status: "verified" });
    expect(result.value.changes.map(({ kind }) => kind)).toEqual(["table-column-insert"]);
    expect(await projectView(result.value.buffer, "final")).toEqual(
      await projectView(target, "final"),
    );
    expect(await projectView(result.value.buffer, "original")).toEqual(
      await projectView(base, "final"),
    );
  });

  test("delete_table_column: physical-cell anchors resolve the matching grid column", async () => {
    const base = await buildColumnTableDocx([
      [{ text: "Account" }, { text: "Currency" }, { text: "Status" }],
      [{ text: "Fees" }, { text: "EUR" }, { text: "Open" }],
    ]);
    const target = await buildColumnTableDocx([
      [{ text: "Account" }, { text: "Status" }],
      [{ text: "Fees" }, { text: "Open" }],
    ]);

    const result = await compareDocx(base, target, OPTIONS);
    if (result.isErr()) {
      throw result.error;
    }
    expect(result.value.verification).toEqual({ status: "verified" });
    expect(result.value.changes.map(({ kind }) => kind)).toEqual(["table-column-delete"]);
    expect(await projectView(result.value.buffer, "final")).toEqual(
      await projectView(target, "final"),
    );
    expect(await projectView(result.value.buffer, "original")).toEqual(
      await projectView(base, "final"),
    );
  });

  test("multiple table columns round-trip in one comparison", async () => {
    const base = await buildColumnTableDocx([
      [{ text: "Account" }, { text: "Status" }],
      [{ text: "Fees" }, { text: "Open" }],
    ]);
    const target = await buildColumnTableDocx([
      [{ text: "Account" }, { text: "Currency" }, { text: "Region" }, { text: "Status" }],
      [{ text: "Fees" }, { text: "EUR" }, { text: "EMEA" }, { text: "Open" }],
    ]);

    const result = await compareDocx(base, target, OPTIONS);
    if (result.isErr()) {
      throw result.error;
    }
    expect(result.value.verification).toEqual({ status: "verified" });
    expect(result.value.changes.map(({ kind }) => kind)).toEqual([
      "table-column-insert",
      "table-column-insert",
    ]);
    expect(await projectView(result.value.buffer, "final")).toEqual(
      await projectView(target, "final"),
    );
    expect(await projectView(result.value.buffer, "original")).toEqual(
      await projectView(base, "final"),
    );
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

  test("format_only_strike: striking a phrase round-trips as tracked formatting", async () => {
    const blockIndex = wordyBlockIndex(PROSE_BLOCKS, 4);
    const target = await applyEditScript(PROSE_BASE, [
      {
        type: "formatRange",
        blockIndex,
        startOffset: 0,
        endOffset: 4,
        formatting: { strike: true },
      },
    ]);
    if (target.isErr()) {
      throw target.error;
    }

    const result = await compareDocx(PROSE_BASE, target.value.buffer, OPTIONS);
    if (result.isErr()) {
      throw result.error;
    }

    expect(result.value.changes.map(({ kind }) => kind)).toEqual(["format"]);
    expect(result.value.verification).toEqual({ status: "verified" });
    const accepted = await FolioDocxReviewer.fromBuffer(result.value.buffer);
    const acceptedBlock = accepted.readReviewedStory({ view: "final" })?.snapshot.blocks[
      blockIndex
    ];
    const rejectedBlock = accepted.readReviewedStory({ view: "original" })?.snapshot.blocks[
      blockIndex
    ];
    expect(acceptedBlock?.previewRuns?.at(0)?.strike).toBe(true);
    expect(rejectedBlock?.previewRuns?.at(0)?.strike).not.toBe(true);
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

  test("insert_paragraph_beside_list_item: the new paragraph is not a list item", async () => {
    // An insertion inherits the anchor's numbering unless it says otherwise,
    // and the anchor is whichever block happened to sit next to it. Saying
    // nothing left every paragraph added beside a list as a further item of
    // that list, which the self-check then refused.
    const withParagraph = [
      ...NUMBERED_LIST_ITEMS.slice(0, 2),
      { level: null, text: "The following item restates the delivery duty." },
      ...NUMBERED_LIST_ITEMS.slice(2),
    ];
    const target = await buildNumberedListDocx(withParagraph);
    const result = await compareDocx(LIST_BASE, target, OPTIONS);
    if (result.isErr()) {
      throw result.error;
    }
    expect(result.value.changes.map(({ kind }) => kind)).toEqual(["insert"]);

    const reviewer = await FolioDocxReviewer.fromBuffer(result.value.buffer);
    reviewer.resolveReviewedStory({ view: "final" });
    expect(reviewer.getContent().map(({ listLevel }) => listLevel ?? null)).toEqual(
      withParagraph.map(({ level }) => level),
    );
  });

  test("delete_paragraph: the paragraph MARK is deleted with the words", async () => {
    // Deleting only the runs leaves the mark, so a consumer that accepts the
    // redline is left with a blank line where the paragraph was. A deleted
    // paragraph carries `w:pPr/w:rPr/w:del` too, and that mark is what makes
    // the paragraph itself go away.
    const clauses = [
      "The parties agree as set out below.",
      "This clause is withdrawn by the target.",
      "This agreement is governed by the stated law.",
    ] as const;
    const base = await buildBodySequenceDocx(
      clauses.map((text) => ({ kind: "paragraph", text }) as const),
    );
    const target = await buildBodySequenceDocx(
      [clauses[0], clauses[2]].map((text) => ({ kind: "paragraph", text }) as const),
    );

    const result = await compareDocx(base, target, OPTIONS);
    if (result.isErr()) {
      throw result.error;
    }
    expect(result.value.changes.map(({ kind }) => kind)).toEqual(["delete"]);
    const deleted = paragraphsOf(await documentPartOf(result.value.buffer)).filter((paragraph) =>
      paragraph.includes("<w:delText"),
    );
    expect(deleted).toHaveLength(1);
    expect(deleted.at(0)).toMatch(/<w:pPr>[\s\S]*<w:rPr>[\s\S]*<w:del\s/u);

    expect(await projectView(result.value.buffer, "final")).toEqual(
      await projectView(target, "final"),
    );
    expect(await projectView(result.value.buffer, "original")).toEqual(
      await projectView(base, "final"),
    );
  });

  test("insert_paragraph: the paragraph MARK is inserted with the words", async () => {
    const clauses = [
      "The parties agree as set out below.",
      "A clause the target adds between the two.",
      "This agreement is governed by the stated law.",
    ] as const;
    const base = await buildBodySequenceDocx(
      [clauses[0], clauses[2]].map((text) => ({ kind: "paragraph", text }) as const),
    );
    const target = await buildBodySequenceDocx(
      clauses.map((text) => ({ kind: "paragraph", text }) as const),
    );

    const result = await compareDocx(base, target, OPTIONS);
    if (result.isErr()) {
      throw result.error;
    }
    expect(result.value.changes.map(({ kind }) => kind)).toEqual(["insert"]);
    const inserted = paragraphsOf(await documentPartOf(result.value.buffer)).filter(
      (paragraph) => paragraph.includes("<w:ins ") && !paragraph.includes("<w:delText"),
    );
    expect(inserted).toHaveLength(1);
    expect(inserted.at(0)).toMatch(/<w:pPr>[\s\S]*<w:rPr>[\s\S]*<w:ins\s/u);

    expect(await projectView(result.value.buffer, "final")).toEqual(
      await projectView(target, "final"),
    );
    expect(await projectView(result.value.buffer, "original")).toEqual(
      await projectView(base, "final"),
    );
  });

  test("delete_last_paragraph_of_a_cell: no mark, because nothing follows it", async () => {
    // A paragraph-mark revision joins the paragraph with the one after it, so
    // on the last paragraph of a cell it could not do what it says: resolving
    // it would merge the cells rather than the paragraphs.
    const base = await buildBodySequenceDocx([
      { kind: "paragraph", text: "The schedule below records the agreed fees." },
      { kind: "table", rows: [["Service", "Fee"]] },
      { kind: "paragraph", text: "This agreement is governed by the stated law." },
    ]);
    const target = await buildBodySequenceDocx([
      { kind: "paragraph", text: "The schedule below records the agreed fees." },
      { kind: "table", rows: [["Service", ""]] },
      { kind: "paragraph", text: "This agreement is governed by the stated law." },
    ]);

    const result = await compareDocx(base, target, OPTIONS);
    if (result.isErr()) {
      throw result.error;
    }
    const deleted = paragraphsOf(await documentPartOf(result.value.buffer)).filter((paragraph) =>
      paragraph.includes("<w:delText"),
    );
    expect(deleted).toHaveLength(1);
    expect(deleted.at(0)).not.toMatch(/<w:pPr>[\s\S]*<w:rPr>[\s\S]*<w:del\s/u);
  });

  test("delete_table_before_terminal_carrier: final paragraph mark remains untracked", async () => {
    // Microsoft Word keeps a body carrier after a terminal table. A deletion
    // on that paragraph mark cannot resolve because there is no next body
    // paragraph to join it to.
    const keptTable = { kind: "table", rows: [["Kept row"]] } as const;
    const base = await buildBodySequenceDocx([
      keptTable,
      { kind: "paragraph", text: "" },
      { kind: "table", rows: [["Removed row"]] },
    ]);
    const target = await buildBodySequenceDocx([keptTable]);

    const result = await compareDocx(base, target, OPTIONS);
    if (result.isErr()) {
      throw result.error;
    }

    expect(result.value.changes.map(({ kind }) => kind)).toEqual(["delete", "table-delete"]);
    const terminalParagraph = paragraphsOf(await documentPartOf(result.value.buffer)).at(-1);
    expect(terminalParagraph).toBeDefined();
    expect(terminalParagraph).not.toMatch(/<w:pPr>[\s\S]*<w:rPr>[\s\S]*<w:del\b/u);
    expect(await projectView(result.value.buffer, "final")).toEqual(
      await projectView(target, "final"),
    );
    expect(await projectView(result.value.buffer, "original")).toEqual(
      await projectView(base, "final"),
    );
  });

  test("unrepresentable_difference: refused by default, emitted and named on request", async () => {
    // Every column is empty, so there is no evidence for which of the three
    // target columns is new. Guessing would produce a plausible but misleading
    // structural change; the conservative alignment leaves it unrepresented.
    const base = await buildBodySequenceDocx([
      { kind: "paragraph", text: "The schedule below records the agreed fees." },
      { kind: "table", rows: [["", ""]] },
      { kind: "paragraph", text: "This agreement is governed by the stated law." },
    ]);
    const target = await buildBodySequenceDocx([
      { kind: "paragraph", text: "The schedule below records the agreed fees." },
      { kind: "table", rows: [["", "", ""]] },
      { kind: "paragraph", text: "This agreement is governed by the stated law." },
    ]);

    const refused = await compareDocx(base, target, OPTIONS);
    expect(refused.isErr()).toBe(true);
    if (refused.isErr()) {
      const error = refused.error;
      expect(error._tag).toBe("CompareDocxRoundTripError");
      if (error._tag === "CompareDocxRoundTripError") {
        expect(error.invariant).toBe("accept-reproduces-target");
        expect(error.cause).toBe("container");
        expect(error.failures.length).toBeGreaterThan(0);
        // Structural facts only: nothing a document said.
        expect(error.failures.at(0)?.detail).not.toContain("schedule below");
      }
    }

    const emitted = await compareDocx(base, target, { ...OPTIONS, onUnverified: "emit" });
    if (emitted.isErr()) {
      throw emitted.error;
    }
    expect(emitted.value.buffer.byteLength).toBeGreaterThan(0);
    expect(emitted.value.changes.length).toBeGreaterThan(0);
    expect(emitted.value.verification.status).toBe("unverified");
    if (emitted.value.verification.status === "unverified") {
      expect(emitted.value.verification.failures.map(({ cause }) => cause)).toContain("container");
    }
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
