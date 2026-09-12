/**
 * Labelled single mutations, one probe each.
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
import type { FolioAIBlock, FolioAIInlineBooleanProperty } from "../ai-edits/types";
import { createDocx } from "../docx/rezip";
import type { Table, TableCell, TextFormatting } from "../types/document";
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

const partOf = async (buffer: ArrayBuffer, name: string): Promise<string> =>
  (await (await JSZip.loadAsync(buffer)).file(name)?.async("string")) ?? "";

type NumberingFormatOverride = { level: number; format: string };

const withNumberingFormats = async (
  buffer: ArrayBuffer,
  overrides: readonly NumberingFormatOverride[],
): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(buffer);
  const entry = zip.file("word/numbering.xml");
  if (!entry) {
    throw new Error("The numbered-list fixture has no numbering part.");
  }
  let xml = await entry.async("string");
  for (const { level, format } of overrides) {
    const levelFormat = new RegExp(
      `(<w:lvl w:ilvl="${String(level)}">[\\s\\S]*?<w:numFmt w:val=")[^"]+("/>)`,
      "u",
    );
    const replaced = xml.replace(levelFormat, (_match, before: string, after: string) =>
      [before, format, after].join(""),
    );
    if (replaced === xml) {
      throw new Error(`The numbered-list fixture has no level ${String(level)} format.`);
    }
    xml = replaced;
  }
  zip.file("word/numbering.xml", xml);
  return await zip.generateAsync({ type: "arraybuffer" });
};

const documentPartOf = async (buffer: ArrayBuffer): Promise<string> =>
  await partOf(buffer, "word/document.xml");

const DIRECT_BOOLEAN_STATES = ["absent", "on", "off"] as const;
type DirectBooleanState = (typeof DIRECT_BOOLEAN_STATES)[number];

const INLINE_BOOLEAN_PROPERTIES = [
  "bold",
  "italic",
  "underline",
  "strike",
] as const satisfies readonly FolioAIInlineBooleanProperty[];

const booleanTextFormatting = (
  property: FolioAIInlineBooleanProperty,
  value: boolean,
): TextFormatting => {
  switch (property) {
    case "bold":
      return { bold: value };
    case "italic":
      return { italic: value };
    case "underline":
      return { underline: { style: value ? "single" : "none" } };
    case "strike":
      return { strike: value };
  }
};

const canonicalBooleanFormattingValue = (property: FolioAIInlineBooleanProperty, value: boolean) =>
  property === "underline"
    ? {
        type: "object" as const,
        entries: [{ key: "style", value: value ? "single" : "none" }],
      }
    : value;

const directBooleanPresence = (
  property: FolioAIInlineBooleanProperty,
  state: DirectBooleanState,
) =>
  state === "absent"
    ? ({ type: "absent" } as const)
    : ({
        type: "present",
        value: canonicalBooleanFormattingValue(property, state === "on"),
      } as const);

const effectiveBooleanPresence = (property: FolioAIInlineBooleanProperty, value: boolean) =>
  ({
    type: "present",
    value: canonicalBooleanFormattingValue(property, value),
  }) as const;

const effectiveBooleanValue = (state: DirectBooleanState, inherited: boolean): boolean => {
  if (state === "absent") {
    return inherited;
  }
  return state === "on";
};

const rgbPropertyValue = (rgb: string) => ({
  type: "object" as const,
  entries: [{ key: "rgb", value: rgb }],
});

const fontFamilyPropertyValue = (fontFamily: string) => ({
  type: "object" as const,
  entries: [
    { key: "ascii", value: fontFamily },
    { key: "hAnsi", value: fontFamily },
  ],
});

const buildBooleanFormattingDocx = (
  property: FolioAIInlineBooleanProperty,
  inherited: boolean,
  direct: DirectBooleanState,
): Promise<ArrayBuffer> => {
  const document = createEmptyDocument();
  const docDefaults = document.package.styles?.docDefaults;
  document.package.styles = {
    ...document.package.styles,
    docDefaults: {
      ...docDefaults,
      rPr: { ...docDefaults?.rPr, ...booleanTextFormatting(property, inherited) },
    },
  };
  document.package.document.content = [
    {
      type: "paragraph",
      paraId: "A10000B0",
      content: [
        {
          type: "run",
          ...(direct !== "absent" && {
            formatting: booleanTextFormatting(property, direct === "on"),
          }),
          content: [{ type: "text", text: "Boolean formatting" }],
        },
      ],
    },
  ];
  return createDocx(document);
};

type BooleanFormattingProjection = {
  direct: DirectBooleanState;
  effective: boolean;
};

const directBooleanState = (value: boolean | undefined): DirectBooleanState => {
  if (value === undefined) {
    return "absent";
  }
  return value ? "on" : "off";
};

const booleanFormattingValue = (
  formatting: TextFormatting | undefined,
  property: FolioAIInlineBooleanProperty,
): boolean | undefined => {
  if (property === "underline") {
    const underline = formatting?.underline;
    return underline === undefined ? undefined : underline.style !== "none";
  }
  return formatting?.[property];
};

const booleanFormattingProjection = async (
  buffer: ArrayBuffer,
  property: FolioAIInlineBooleanProperty,
): Promise<BooleanFormattingProjection> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(buffer);
  const run = reviewer.snapshot().blocks.at(0)?.previewRuns?.at(0);
  return {
    direct: directBooleanState(booleanFormattingValue(run?.authoredFormatting, property)),
    effective: booleanFormattingValue(run?.effectiveFormatting, property) === true,
  };
};

const withoutTerminalBodyParagraph = async (buffer: ArrayBuffer): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(buffer);
  const documentEntry = zip.file("word/document.xml");
  if (!documentEntry) {
    throw new Error("The fixture has no main document part.");
  }
  const documentXml = await documentEntry.async("string");
  const withoutCarrier = documentXml.replace(/<w:p\b[^>]*><\/w:p>(?=<w:sectPr>)/u, "");
  if (withoutCarrier === documentXml) {
    throw new Error("The fixture has no terminal empty body paragraph.");
  }
  zip.file("word/document.xml", withoutCarrier);
  return await zip.generateAsync({ type: "arraybuffer" });
};

/** Every `w:p` element of a document part, as raw XML. */
const paragraphsOf = (documentXml: string): string[] =>
  documentXml.match(/<w:p[ >][\s\S]*?<\/w:p>/gu) ?? [];

/**
 * Whether the paragraph's OWN mark carries a deletion. `w:pPr` is a paragraph's
 * first child, so reading only that far keeps a deleted RUN's own `w:rPr` out
 * of the answer.
 */
const markIsDeleted = (paragraph: string): boolean => {
  const properties = /^<w:p\b[^>]*><w:pPr>([\s\S]*?)<\/w:pPr>/u.exec(paragraph)?.[1] ?? "";
  return /<w:rPr>[\s\S]*?<w:(?:del|moveFrom)\b/u.test(properties);
};

/** Every paragraph of a part whose own mark carries a deletion, in order. */
const marksDeletedIn = (partXml: string): string[] =>
  paragraphsOf(partXml).filter((paragraph) => markIsDeleted(paragraph));

const markIsInserted = (paragraph: string): boolean => {
  const properties = /^<w:p\b[^>]*><w:pPr>([\s\S]*?)<\/w:pPr>/u.exec(paragraph)?.[1] ?? "";
  return /<w:rPr>[\s\S]*?<w:(?:ins|moveTo)\b/u.test(properties);
};

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

/** The same projection with the paragraph style, for a properties assertion. */
const styledView = async (
  buffer: ArrayBuffer,
  view: "original" | "final",
): Promise<(BlockProjection & { styleId: string | null })[]> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(buffer);
  const story = reviewer.readReviewedStory({ view });
  return (story?.snapshot.blocks ?? []).map((block) => ({
    text: block.text,
    table: block.table ?? null,
    styleId: block.styleId ?? null,
  }));
};

type ProbeOutcome = {
  buffer: ArrayBuffer;
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

  return {
    buffer: result.value.buffer,
    changes: result.value.changes,
    kinds: result.value.changes.map(({ kind }) => kind),
  };
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

  for (const property of INLINE_BOOLEAN_PROPERTIES) {
    for (const inherited of [false, true]) {
      for (const baseDirect of DIRECT_BOOLEAN_STATES) {
        for (const targetDirect of DIRECT_BOOLEAN_STATES) {
          if (baseDirect === targetDirect) {
            continue;
          }
          test(`format_only_${property}: inherited ${String(inherited)}, direct ${baseDirect} -> ${targetDirect}`, async () => {
            const base = await buildBooleanFormattingDocx(property, inherited, baseDirect);
            const target = await buildBooleanFormattingDocx(property, inherited, targetDirect);
            const result = await compareDocx(base, target, OPTIONS);
            if (result.isErr()) {
              throw result.error;
            }

            expect(result.value.verification).toEqual({ status: "verified" });
            const baseEffective = effectiveBooleanValue(baseDirect, inherited);
            const revisedEffective = effectiveBooleanValue(targetDirect, inherited);
            expect(result.value.changes).toEqual([
              expect.objectContaining({
                kind: "format",
                ranges: [
                  {
                    startOffset: 0,
                    endOffset: "Boolean formatting".length,
                    formatting: {
                      authored: [
                        {
                          key: property,
                          base: directBooleanPresence(property, baseDirect),
                          revised: directBooleanPresence(property, targetDirect),
                        },
                      ],
                      effective:
                        baseEffective === revisedEffective
                          ? []
                          : [
                              {
                                key: property,
                                base: effectiveBooleanPresence(property, baseEffective),
                                revised: effectiveBooleanPresence(property, revisedEffective),
                              },
                            ],
                    },
                  },
                ],
              }),
            ]);

            const accepting = await FolioDocxReviewer.fromBuffer(result.value.buffer);
            accepting.acceptAll();
            const accepted = await accepting.toBuffer();
            const rejecting = await FolioDocxReviewer.fromBuffer(result.value.buffer);
            rejecting.rejectAll();
            const rejected = await rejecting.toBuffer();

            expect((await FolioDocxReviewer.fromBuffer(accepted)).getChanges()).toEqual([]);
            expect((await FolioDocxReviewer.fromBuffer(rejected)).getChanges()).toEqual([]);
            expect(await booleanFormattingProjection(accepted, property)).toEqual(
              await booleanFormattingProjection(target, property),
            );
            expect(await booleanFormattingProjection(rejected, property)).toEqual(
              await booleanFormattingProjection(base, property),
            );
            expect(await documentPartOf(accepted)).not.toContain("<w:rPrChange ");
            expect(await documentPartOf(rejected)).not.toContain("<w:rPrChange ");
          });
        }
      }
    }
  }

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
    expect(acceptedBlock?.previewRuns?.at(0)?.effectiveFormatting?.strike).toBe(true);
    expect(rejectedBlock?.previewRuns?.at(0)?.effectiveFormatting?.strike).not.toBe(true);
  });

  test("format_only_font: changing face, half-point size, and color is one format", async () => {
    const blockIndex = wordyBlockIndex(PROSE_BLOCKS, 4);
    const { buffer, changes, kinds } = await probe(PROSE_BASE, [
      {
        type: "formatRange",
        blockIndex,
        startOffset: 0,
        endOffset: 4,
        formatting: { fontFamily: "Georgia", fontSizePt: 10.5, color: "C00000" },
      },
    ]);
    expect(kinds).toEqual(["format"]);
    const change = changes.at(0);
    expect(change?.kind === "format" ? change.ranges : []).toEqual([
      {
        startOffset: 0,
        endOffset: 4,
        formatting: {
          authored: [
            {
              key: "color",
              base: { type: "absent" },
              revised: { type: "present", value: rgbPropertyValue("C00000") },
            },
            {
              key: "fontFamily",
              base: { type: "absent" },
              revised: { type: "present", value: fontFamilyPropertyValue("Georgia") },
            },
            {
              key: "fontSize",
              base: { type: "absent" },
              revised: { type: "present", value: 21 },
            },
          ],
          effective: [
            {
              key: "color",
              base: { type: "absent" },
              revised: { type: "present", value: rgbPropertyValue("C00000") },
            },
            {
              key: "fontFamily",
              base: { type: "present", value: fontFamilyPropertyValue("Calibri") },
              revised: { type: "present", value: fontFamilyPropertyValue("Georgia") },
            },
            {
              key: "fontSize",
              base: { type: "present", value: 22 },
              revised: { type: "present", value: 21 },
            },
          ],
        },
      },
    ]);
    const reviewed = await FolioDocxReviewer.fromBuffer(buffer);
    expect(
      reviewed
        .readReviewedStory({ view: "final" })
        ?.snapshot.blocks[blockIndex]?.previewRuns?.at(0),
    ).toMatchObject({
      effectiveFormatting: {
        color: { rgb: "C00000" },
        fontFamily: { ascii: "Georgia", hAnsi: "Georgia" },
        fontSize: 21,
      },
      authoredFormatting: {
        color: { rgb: "C00000" },
        fontFamily: { ascii: "Georgia", hAnsi: "Georgia" },
        fontSize: 21,
      },
    });
  });

  test("format_only_font_clear: clearing direct font properties is one format", async () => {
    const blockIndex = wordyBlockIndex(PROSE_BLOCKS, 4);
    const formatted = await applyEditScript(PROSE_BASE, [
      {
        type: "formatRange",
        blockIndex,
        startOffset: 0,
        endOffset: 4,
        formatting: { fontFamily: "Georgia", fontSizePt: 10.5, color: "C00000" },
      },
    ]);
    if (formatted.isErr()) {
      throw formatted.error;
    }

    const { buffer, changes, kinds } = await probe(formatted.value.buffer, [
      {
        type: "formatRange",
        blockIndex,
        startOffset: 0,
        endOffset: 4,
        formatting: { fontFamily: null, fontSizePt: null, color: null },
      },
    ]);
    expect(kinds).toEqual(["format"]);
    expect(changes.at(0)?.kind === "format" ? changes.at(0)?.ranges : []).toEqual([
      {
        startOffset: 0,
        endOffset: 4,
        formatting: {
          authored: [
            {
              key: "color",
              base: { type: "present", value: rgbPropertyValue("C00000") },
              revised: { type: "absent" },
            },
            {
              key: "fontFamily",
              base: { type: "present", value: fontFamilyPropertyValue("Georgia") },
              revised: { type: "absent" },
            },
            {
              key: "fontSize",
              base: { type: "present", value: 21 },
              revised: { type: "absent" },
            },
          ],
          effective: [
            {
              key: "color",
              base: { type: "present", value: rgbPropertyValue("C00000") },
              revised: { type: "absent" },
            },
            {
              key: "fontFamily",
              base: { type: "present", value: fontFamilyPropertyValue("Georgia") },
              revised: { type: "present", value: fontFamilyPropertyValue("Calibri") },
            },
            {
              key: "fontSize",
              base: { type: "present", value: 21 },
              revised: { type: "present", value: 22 },
            },
          ],
        },
      },
    ]);
    const reviewed = await FolioDocxReviewer.fromBuffer(buffer);
    expect(
      reviewed.readReviewedStory({ view: "final" })?.snapshot.blocks[blockIndex]?.previewRuns?.at(0)
        ?.authoredFormatting,
    ).toBeUndefined();
    expect(
      reviewed
        .readReviewedStory({ view: "original" })
        ?.snapshot.blocks[blockIndex]?.previewRuns?.at(0),
    ).toMatchObject({
      effectiveFormatting: {
        color: { rgb: "C00000" },
        fontFamily: { ascii: "Georgia", hAnsi: "Georgia" },
        fontSize: 21,
      },
      authoredFormatting: {
        color: { rgb: "C00000" },
        fontFamily: { ascii: "Georgia", hAnsi: "Georgia" },
        fontSize: 21,
      },
    });
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
    const result = await compareDocx(LIST_BASE, roman, { ...OPTIONS, mode: "bestEffort" });
    if (result.isErr()) {
      throw result.error;
    }
    const kinds = result.value.changes.map(({ kind }) => kind);
    expect(new Set(kinds)).toEqual(new Set(["numbering"]));
    const [change] = result.value.changes;
    expect(change?.kind === "numbering" && change.before?.format).toBe("decimal");
    expect(change?.kind === "numbering" && change.after?.format).toBe("lowerRoman");
    expect(result.value.verification.status).toBe("unverified");
    if (result.value.verification.status === "unverified") {
      expect(
        result.value.verification.failures.filter(({ scope }) => scope.type === "package"),
      ).toEqual([
        {
          invariant: "accept-reproduces-target",
          cause: "unsupported",
          scope: { type: "package" },
          detail: "the numbering-definition difference has no proved tracked-document instruction",
        },
      ]);
    }
  });

  test("numbering definitions: only changed levels referenced by either document are reported", async () => {
    const changed = await withNumberingFormats(LIST_BASE, [
      { level: 0, format: "upperRoman" },
      { level: 1, format: "upperLetter" },
      { level: 2, format: "ordinal" },
    ]);
    const result = await compareDocx(LIST_BASE, changed, OPTIONS);
    if (result.isErr()) {
      throw result.error;
    }

    const numberingChanges = result.value.changes.filter(({ kind }) => kind === "numbering");
    expect(numberingChanges.map(({ level }) => level)).toEqual([0, 1]);
  });

  test("numbering definitions: a changed level newly referenced by the target is reported", async () => {
    const items = withItemDemoted(NUMBERED_LIST_ITEMS, 1);
    const changed = await withNumberingFormats(await buildNumberedListDocx(items), [
      { level: 2, format: "upperRoman" },
    ]);
    const result = await compareDocx(LIST_BASE, changed, OPTIONS);
    if (result.isErr()) {
      throw result.error;
    }

    const numberingChanges = result.value.changes.filter(({ kind }) => kind === "numbering");
    expect(numberingChanges.map(({ level }) => level)).toEqual([2]);
  });

  test("numbering definitions: a changed level referenced only by the base is reported", async () => {
    const baseItems = withItemDemoted(NUMBERED_LIST_ITEMS, 1);
    const base = await buildNumberedListDocx(baseItems);
    const changed = await withNumberingFormats(LIST_BASE, [{ level: 2, format: "upperRoman" }]);
    const result = await compareDocx(base, changed, OPTIONS);
    if (result.isErr()) {
      throw result.error;
    }

    const numberingChanges = result.value.changes.filter(({ kind }) => kind === "numbering");
    expect(numberingChanges.map(({ level }) => level)).toEqual([2]);
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
    // A body may not end with a table, so a package that does carries a
    // paragraph after it. A deletion on that paragraph's mark cannot resolve:
    // there is no next body paragraph to join it to.
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

  test("replace_terminal_table: the base carrier owns the deletion before the target table", async () => {
    const base = await buildBodySequenceDocx([
      { kind: "table", rows: [["Source terminal table"]] },
      { kind: "paragraph", text: "" },
    ]);
    const target = await withoutTerminalBodyParagraph(
      await buildBodySequenceDocx([
        {
          kind: "table",
          rows: [[{ content: "Target terminal table", gridSpan: 2 }]],
        },
      ]),
    );
    expect(await projectView(base, "final")).toHaveLength(2);
    expect(await projectView(target, "final")).toHaveLength(1);
    expect(await documentPartOf(target)).toMatch(/<\/w:tbl><w:sectPr>/u);

    const result = await compareDocx(base, target, OPTIONS);
    if (result.isErr()) {
      throw result.error;
    }

    expect(result.value.changes.map(({ kind }) => kind)).toEqual([
      "table-delete",
      "table-insert",
      "delete",
    ]);
    const pendingXml = await documentPartOf(result.value.buffer);
    expect(pendingXml).toMatch(
      /<w:tbl>[\s\S]*?<w:trPr><w:del\b[^>]*\/><\/w:trPr>[\s\S]*?Source terminal table/u,
    );
    expect(pendingXml).toMatch(
      /Source terminal table[\s\S]*?<\/w:tbl><w:p\b[^>]*><w:pPr><w:rPr><w:del\b[\s\S]*?<\/w:p><w:tbl>[\s\S]*?Target terminal table/u,
    );
    expect(pendingXml).toMatch(
      /<w:tbl>[\s\S]*?<w:trPr><w:ins\b[^>]*\/><\/w:trPr>[\s\S]*?Target terminal table/u,
    );
    expect(pendingXml).toMatch(/<\/w:tbl><w:sectPr>/u);

    const acceptedReviewer = await FolioDocxReviewer.fromBuffer(result.value.buffer);
    acceptedReviewer.acceptAll();
    const accepted = await acceptedReviewer.toBuffer();
    const rejectedReviewer = await FolioDocxReviewer.fromBuffer(result.value.buffer);
    rejectedReviewer.rejectAll();
    const rejected = await rejectedReviewer.toBuffer();

    const acceptedReopened = await FolioDocxReviewer.fromBuffer(accepted);
    const rejectedReopened = await FolioDocxReviewer.fromBuffer(rejected);
    expect(acceptedReopened.readReviewedStory({ view: "current-markup" })?.changes).toEqual([]);
    expect(rejectedReopened.readReviewedStory({ view: "current-markup" })?.changes).toEqual([]);
    expect(await projectView(accepted, "final")).toEqual(await projectView(target, "final"));
    expect(await projectView(rejected, "final")).toEqual(await projectView(base, "final"));
    expect(await documentPartOf(accepted)).toMatch(/<\/w:tbl><w:sectPr>/u);
    expect(await documentPartOf(rejected)).toMatch(/<\/w:p><w:sectPr>/u);
  });

  /**
   * Four paragraphs become one: the first is edited and the three after it are
   * removed. The last of them ends the body, so its mark cannot be deleted —
   * nothing follows it to merge into. The chain therefore starts at the last
   * SURVIVING paragraph: the marks of the first three go, the fourth keeps its
   * own, and the merged paragraph lands on it.
   */
  const TRAILING_CLAUSES = [
    "Alpha clause states the agreed position.",
    "Bravo clause states the agreed position.",
    "Charlie clause states the agreed position.",
    "Delta clause states the agreed position.",
  ] as const;
  const REVISED_CLAUSE = "Alpha clause states the revised position.";

  test("delete_trailing_paragraphs: the body's final mark stays, the chain moves back", async () => {
    const base = await buildBodySequenceDocx(
      TRAILING_CLAUSES.map((text) => ({ kind: "paragraph", text }) as const),
    );
    const target = await buildBodySequenceDocx([{ kind: "paragraph", text: REVISED_CLAUSE }]);

    const result = await compareDocx(base, target, OPTIONS);
    if (result.isErr()) {
      throw result.error;
    }
    expect(result.value.changes.map(({ kind }) => kind)).toEqual([
      "replace",
      "delete",
      "delete",
      "delete",
    ]);

    const paragraphs = paragraphsOf(await documentPartOf(result.value.buffer));
    expect(paragraphs).toHaveLength(4);
    expect(paragraphs.map((paragraph) => markIsDeleted(paragraph))).toEqual([
      true,
      true,
      true,
      false,
    ]);

    expect(await projectView(result.value.buffer, "final")).toEqual(
      await projectView(target, "final"),
    );
    expect(await projectView(result.value.buffer, "original")).toEqual(
      await projectView(base, "final"),
    );
  });

  test("delete_trailing_paragraphs_in_a_cell: the cell's final mark stays", async () => {
    const cell = (texts: readonly string[]) =>
      texts.map((text) => ({ kind: "paragraph", text }) as const);
    const base = await buildBodySequenceDocx([
      { kind: "paragraph", text: "The schedule below records the agreed fees." },
      { kind: "table", rows: [[cell(TRAILING_CLAUSES), "Fee"]] },
      { kind: "paragraph", text: "This agreement is governed by the stated law." },
    ]);
    const target = await buildBodySequenceDocx([
      { kind: "paragraph", text: "The schedule below records the agreed fees." },
      { kind: "table", rows: [[cell([REVISED_CLAUSE]), "Fee"]] },
      { kind: "paragraph", text: "This agreement is governed by the stated law." },
    ]);

    const result = await compareDocx(base, target, OPTIONS);
    if (result.isErr()) {
      throw result.error;
    }
    const documentXml = await documentPartOf(result.value.buffer);
    // Three marks go — the edited paragraph's and the two after it — and the
    // cell's fourth paragraph keeps its own.
    expect(marksDeletedIn(documentXml)).toHaveLength(3);
    const cellParagraphs = paragraphsOf(documentXml).filter((paragraph) =>
      paragraph.includes("clause states"),
    );
    expect(cellParagraphs).toHaveLength(4);
    expect(markIsDeleted(cellParagraphs.at(-1) ?? "")).toBe(false);

    expect(await projectView(result.value.buffer, "final")).toEqual(
      await projectView(target, "final"),
    );
    expect(await projectView(result.value.buffer, "original")).toEqual(
      await projectView(base, "final"),
    );
  });

  test("delete_trailing_paragraphs_in_a_header: the header story keeps its final mark", async () => {
    const body = [{ kind: "paragraph", text: "The parties agree as set out below." }] as const;
    const base = await buildBodySequenceDocx(body, {
      header: TRAILING_CLAUSES.map((text) => ({ kind: "paragraph", text }) as const),
    });
    const target = await buildBodySequenceDocx(body, {
      header: [{ kind: "paragraph", text: REVISED_CLAUSE }],
    });

    const result = await compareDocx(base, target, OPTIONS);
    if (result.isErr()) {
      throw result.error;
    }
    expect(result.value.unsupported).toEqual([]);
    const headerParagraphs = paragraphsOf(await partOf(result.value.buffer, "word/header1.xml"));
    expect(headerParagraphs).toHaveLength(4);
    expect(headerParagraphs.map((paragraph) => markIsDeleted(paragraph))).toEqual([
      true,
      true,
      true,
      false,
    ]);
  });

  test("trailing_deletion_moves_the_surviving_properties_to_the_carrier", async () => {
    // The merged paragraph ends on the carrier's mark, and a paragraph's
    // properties live on its mark, so the carrier is where the surviving
    // paragraph's style has to end up: as `w:pPrChange`, which is what a
    // rejection reads to put the carrier's own style back.
    const base = await buildBodySequenceDocx([
      { kind: "paragraph", text: TRAILING_CLAUSES[0], styleId: "Heading1" },
      { kind: "paragraph", text: TRAILING_CLAUSES[1] },
      { kind: "paragraph", text: TRAILING_CLAUSES[2] },
      { kind: "paragraph", text: TRAILING_CLAUSES[3] },
    ]);
    const target = await buildBodySequenceDocx([
      { kind: "paragraph", text: REVISED_CLAUSE, styleId: "Heading1" },
    ]);

    const result = await compareDocx(base, target, OPTIONS);
    if (result.isErr()) {
      throw result.error;
    }
    const carrier = paragraphsOf(await documentPartOf(result.value.buffer)).at(-1) ?? "";
    expect(carrier).toContain(`<w:pStyle w:val="Heading1"/>`);
    expect(carrier).toContain("<w:pPrChange ");
    expect(markIsDeleted(carrier)).toBe(false);

    const [accepted, expected, rejected, original] = await Promise.all([
      styledView(result.value.buffer, "final"),
      styledView(target, "final"),
      styledView(result.value.buffer, "original"),
      styledView(base, "final"),
    ]);
    expect(accepted).toEqual(expected);
    expect(rejected).toEqual(original);
  });

  test("closing_paragraph_after_a_table: the target's last words land in the carrier", async () => {
    // A body may not end with a table, so a document that ends in one carries
    // a blank paragraph after it whose mark the format keeps. Nothing of the
    // body sits between that carrier and the table, so no merge chain reaches
    // it: the paragraph the target ends with is written INTO the carrier as
    // inserted runs before its kept mark, and the story ends where the target
    // ends instead of on a blank line the target does not have.
    const base = await buildBodySequenceDocx([
      { kind: "paragraph", text: TRAILING_CLAUSES[0] },
      { kind: "paragraph", text: TRAILING_CLAUSES[1] },
      { kind: "table", rows: [["Service", "Fee"]] },
      { kind: "paragraph", text: "" },
    ]);
    const target = await buildBodySequenceDocx([
      { kind: "paragraph", text: TRAILING_CLAUSES[0] },
      { kind: "paragraph", text: "The agreement closes on the words below." },
    ]);

    const result = await compareDocx(base, target, OPTIONS);
    if (result.isErr()) {
      throw result.error;
    }

    const paragraphs = paragraphsOf(await documentPartOf(result.value.buffer));
    const carrier = paragraphs.at(-1) ?? "";
    expect(markIsDeleted(carrier)).toBe(false);
    expect(carrier).toContain("<w:ins ");
    expect(carrier).toContain("The agreement closes on the words below.");

    expect(await projectView(result.value.buffer, "final")).toEqual(
      await projectView(target, "final"),
    );
    expect(await projectView(result.value.buffer, "original")).toEqual(
      await projectView(base, "final"),
    );
  });

  test("append_paragraphs: the ADDED break lands one paragraph back", async () => {
    // The mirror of the trailing removal. An inserted mark says the break was
    // added, so rejecting it closes the paragraph it ends back over the next
    // one — which the paragraph a container ENDS with does not have. The break
    // therefore sits between the paragraph the run was appended after and the
    // first appended one: that paragraph's mark is the inserted one, and the
    // paragraph the container now ends with takes the free mark it had.
    const base = await buildBodySequenceDocx([
      { kind: "paragraph", text: TRAILING_CLAUSES[0] },
      { kind: "paragraph", text: TRAILING_CLAUSES[1] },
    ]);
    const target = await buildBodySequenceDocx([
      { kind: "paragraph", text: TRAILING_CLAUSES[0] },
      { kind: "paragraph", text: TRAILING_CLAUSES[1] },
      { kind: "paragraph", text: TRAILING_CLAUSES[2] },
      { kind: "paragraph", text: TRAILING_CLAUSES[3] },
    ]);

    const result = await compareDocx(base, target, OPTIONS);
    if (result.isErr()) {
      throw result.error;
    }
    expect(result.value.changes.map(({ kind }) => kind)).toEqual(["insert", "insert"]);

    const paragraphs = paragraphsOf(await documentPartOf(result.value.buffer));
    expect(paragraphs).toHaveLength(4);
    expect(paragraphs.map((paragraph) => markIsInserted(paragraph))).toEqual([
      false,
      true,
      true,
      false,
    ]);

    expect(await projectView(result.value.buffer, "final")).toEqual(
      await projectView(target, "final"),
    );
    expect(await projectView(result.value.buffer, "original")).toEqual(
      await projectView(base, "final"),
    );
  });

  test("ambiguous_column_change: replaces the table instead of guessing a column", async () => {
    // Every column is empty, so there is no evidence for which of the three
    // target columns is new. Guessing would produce a plausible but misleading
    // column edit; replacing the table preserves both reviewed views exactly.
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

    const result = await compareDocx(base, target, OPTIONS);
    if (result.isErr()) {
      throw result.error;
    }
    expect(result.value.verification.status).toBe("verified");
    expect(result.value.changes.map(({ kind }) => kind)).toEqual(["table-delete", "table-insert"]);
    expect(await projectView(result.value.buffer, "final")).toEqual(
      await projectView(target, "final"),
    );
    expect(await projectView(result.value.buffer, "original")).toEqual(
      await projectView(base, "final"),
    );
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
