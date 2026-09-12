/**
 * Properties of {@link compareDocx} over scripted edits of real fixtures.
 *
 * Every case is ground truth by construction: an {@link EditScript} is applied
 * directly to a base document to build the target, so the difference between
 * the two is known before the comparison runs. The properties then pin what the
 * comparison must recover from it.
 *
 * Fixture provenance and licensing: see
 * `../docx/__tests__/__fixtures__/corpus/PROVENANCE.md`.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";
import { readFileSync } from "node:fs";
import path from "node:path";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { FolioDocxReviewer } from "../ai-edits/headless";
import type { FolioAIBlock } from "../ai-edits/types";
import { compareDocx } from "./compare";
import {
  applyEditScript,
  EDIT_SCRIPT_TABLE_ROW_POSITION,
  type EditScript,
  type EditScriptStep,
} from "./scenario";
import type { CompareChange, CompareResult } from "./types";
import { revisedFinalParagraphMarks } from "./verification";

const FIXTURES_DIR = path.join(import.meta.dir, "../docx/__tests__/__fixtures__/corpus");

const readFixture = (filename: string): ArrayBuffer => {
  const bytes = readFileSync(path.join(FIXTURES_DIR, filename));
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
};

/**
 * A synthetic base built from the empty-document fixture, so the corpus is not
 * the only shape under test: a plain run of paragraphs with no styles, no
 * tables, and no Word-authored paraIds.
 */
const buildSyntheticBase = async (): Promise<ArrayBuffer> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(readFixture("upstream-empty.docx"));
  const anchor = reviewer.snapshot();
  // An empty document is one blank paragraph, so its first block is the anchor
  // the synthetic clauses are built on.
  const anchorId = anchor.blocks.at(0)?.id;
  if (anchorId === undefined) {
    throw new Error("The empty fixture offers no anchor to build a synthetic base on.");
  }
  reviewer.applyOperations(
    Array.from({ length: 6 }, (_unused, index) => ({
      id: `synthetic-${String(index)}`,
      type: "insertAfterBlock" as const,
      blockId: anchorId,
      text: `Synthetic clause ${String(index)} sets out the agreed position.`,
    })),
    { mode: "direct" },
  );
  return await reviewer.toBuffer();
};

const PRIOR_AUTHOR = "a previous reviewer";

/**
 * A base that already carries someone else's tracked changes: three edits
 * applied in `"tracked-changes"` mode, so the document holds unresolved
 * insertions and deletions before the comparison ever sees it.
 */
const withPriorRevisions = async (buffer: ArrayBuffer): Promise<ArrayBuffer> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(buffer, { author: PRIOR_AUTHOR });
  const snapshot = reviewer.snapshot();
  reviewer.applyOperations(
    snapshot.blocks.slice(0, 3).map((block, index) => ({
      id: `prior-${String(index)}`,
      type: "replaceBlock" as const,
      blockId: block.id,
      text: `${block.text} As previously amended.`,
    })),
    {
      mode: "tracked-changes",
      snapshot,
      revisionStamp: { date: "2023-01-01T00:00:00.000Z", idSeed: 900 },
    },
  );
  return await reviewer.toBuffer();
};

const REVISION_ELEMENT_ID =
  /<w:(?:ins|del|moveFrom|moveTo|rPrChange|pPrChange|tblPrChange|trPrChange|tcPrChange|cellIns|cellDel|cellMerge)\b[^>]*\bw:id="(\d+)"/gu;

/**
 * Every revision `w:id` the package carries, across every part.
 *
 * A `w:id` is unique package-wide rather than part-wide, and one logical change
 * can serialize as several wrappers — a redline cut around the words that
 * survived, a revision split around a hyperlink — so the ids only stay distinct
 * if something checks them all together.
 */
const revisionIdsInPackage = async (buffer: ArrayBuffer): Promise<string[]> => {
  const zip = await JSZip.loadAsync(buffer);
  const ids: string[] = [];
  for (const [partPath, file] of Object.entries(zip.files)) {
    if (file.dir || !partPath.startsWith("word/") || !partPath.endsWith(".xml")) {
      continue;
    }
    // oxlint-disable-next-line no-await-in-loop -- the parts share one id space, so they are read together
    const xml = await file.async("text");
    for (const [, id] of xml.matchAll(REVISION_ELEMENT_ID)) {
      ids.push(id ?? "");
    }
  }
  return ids;
};

const PARAGRAPH_ID_ATTRIBUTE =
  /\b(?:w|w14|w15|w16cid):(?:paraId|paraIdParent|textId)="([0-9A-Fa-f]{8})"/gu;

/**
 * Every paragraph id and text-revision marker the package carries. They are
 * `ST_LongHexNumber` with a maximum, so the values are 31-bit, and a package
 * carrying a larger one is a package a consumer refuses.
 */
const paragraphIdsInPackage = async (buffer: ArrayBuffer): Promise<string[]> => {
  const zip = await JSZip.loadAsync(buffer);
  const ids: string[] = [];
  for (const [partPath, file] of Object.entries(zip.files)) {
    if (file.dir || !partPath.startsWith("word/") || !partPath.endsWith(".xml")) {
      continue;
    }
    // oxlint-disable-next-line no-await-in-loop -- the parts share one id space, so they are read together
    const xml = await file.async("text");
    for (const [, id] of xml.matchAll(PARAGRAPH_ID_ATTRIBUTE)) {
      ids.push(id ?? "");
    }
  }
  return ids;
};

const authorsOfChanges = async (buffer: ArrayBuffer): Promise<string[]> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(buffer);
  return [...new Set(reviewer.getChanges().map(({ author }) => author))].toSorted();
};

const FIXTURE_FILES = [
  "upstream-styled-content.docx",
  "upstream-with-tables.docx",
  "upstream-complex-styles.docx",
] as const;

const SYNTHETIC_BASE = await buildSyntheticBase();

const BASE_DOCUMENTS: readonly { name: string; buffer: ArrayBuffer }[] = [
  ...FIXTURE_FILES.map((name) => ({ name, buffer: readFixture(name) })),
  { name: "synthetic", buffer: SYNTHETIC_BASE },
];

const OPTIONS = { author: "compare", timestamp: "2024-03-01T00:00:00.000Z" } as const;

const blocksOf = async (buffer: ArrayBuffer): Promise<FolioAIBlock[]> =>
  (await FolioDocxReviewer.fromBuffer(buffer)).getContent();

/** Parsed once at module scope: `describe` bodies run synchronously. */
const BASE_CASES = await Promise.all(
  BASE_DOCUMENTS.map(async ({ name, buffer }) => ({
    name,
    buffer,
    blocks: await blocksOf(buffer),
  })),
);

/**
 * The text-and-structure projection the round-trip properties compare on:
 * every block's text plus where it sits in a table. Two documents that agree
 * on it hold the same content in the same table shape; block ids and revision
 * bookkeeping are deliberately excluded, since a redline necessarily rewrites
 * those.
 */
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

const compareOrThrow = async (base: ArrayBuffer, target: ArrayBuffer): Promise<CompareResult> => {
  const result = await compareDocx(base, target, OPTIONS);
  if (result.isErr()) {
    throw result.error;
  }
  return result.value;
};

/** Every block index of `blocks`, so a generated step always addresses a real block. */
const blockIndexArb = (blocks: readonly FolioAIBlock[]) => fc.nat({ max: blocks.length - 1 });

/**
 * Body-level block indexes only.
 *
 * A relocation is a deletion here and an insertion there, and an insertion
 * anchored inside a table cell cannot land in that cell: no operation places a
 * paragraph in one, so the applier puts it beside the table instead. Moving a
 * cell's paragraph out to the body, or a body paragraph into a cell, is
 * therefore not an edit this vocabulary can express, and generating one tests
 * the engine against a target it never claimed to reach. The property that
 * every scripted difference is representable means scripted differences have
 * to stay inside what the vocabulary covers; the cell case is the documented
 * `container` limitation, measured on real documents rather than asserted here.
 */
const bodyBlockIndexArb = (blocks: readonly FolioAIBlock[]) => {
  const bodyIndexes = blocks.flatMap((block, index) => (block.table ? [] : [index]));
  return bodyIndexes.length === 0 ? null : fc.constantFrom(...bodyIndexes);
};

/**
 * Words a relocated block needs before the move pass will pair its two halves.
 * Mirrors `MOVE_MINIMUM_WORD_COUNT` in `plan.ts`: below the floor a base-only
 * and a target-only block with the same text are not called one relocation,
 * because boilerplate one-liners — and blank paragraphs, which have no words at
 * all — would pair as spurious moves.
 */
const MOVE_MINIMUM_WORD_COUNT = 3;

/** Body-level blocks with enough words for the move pass to pair them. */
const movableBlockIndexArb = (blocks: readonly FolioAIBlock[]) => {
  const indexes = blocks.flatMap((block, index) =>
    !block.table && block.text.trim().split(/\s+/u).length >= MOVE_MINIMUM_WORD_COUNT
      ? [index]
      : [],
  );
  return indexes.length === 0 ? null : fc.constantFrom(...indexes);
};

const wordArb = fc.stringMatching(/^[A-Za-z]{3,9}$/u);
const sentenceArb = fc
  .array(wordArb, { minLength: 3, maxLength: 6 })
  .map((words) => words.join(" "));
const rowRewriteArb = fc
  .array(wordArb, { minLength: 4, maxLength: 6 })
  .map((words) => words.join(" "));

/** A word already present in some block, so `replaceWords` resolves. */
const findableWordArb = (blocks: readonly FolioAIBlock[]) => {
  const candidates: { blockIndex: number; find: string }[] = [];
  blocks.forEach((block, blockIndex) => {
    const word = /[A-Za-z]{3,}/u.exec(block.text);
    if (word) {
      candidates.push({ blockIndex, find: word[0] });
    }
  });
  return candidates.length === 0 ? null : fc.constantFrom(...candidates);
};

const tableBlockIndexArb = (blocks: readonly FolioAIBlock[]) => {
  const indexes = blocks.flatMap((block, index) => (block.table ? [index] : []));
  return indexes.length === 0 ? null : fc.constantFrom(...indexes);
};

/**
 * Blocks in the same table row as `blocks[index]`. An inserted row is
 * generated with exactly this many cell texts: a cell left empty carries no
 * block at all, so a short row would erase the row boundary the comparison
 * reads structure from, and the script would no longer describe what changed.
 */
const rowBlockCount = (blocks: readonly FolioAIBlock[], index: number): number => {
  const table = blocks[index]?.table;
  if (!table) {
    return 1;
  }
  return blocks.filter(
    (block) =>
      block.table?.tableIndex === table.tableIndex && block.table.rowIndex === table.rowIndex,
  ).length;
};

const formattingArb = fc
  .record({
    bold: fc.boolean(),
    italic: fc.boolean(),
    underline: fc.boolean(),
    strike: fc.boolean(),
  })
  .filter(({ bold, italic, underline, strike }) => bold || italic || underline || strike);

/** Steps that only change formatting, for the formatting-only property. */
const formatStepArb = (blocks: readonly FolioAIBlock[]): fc.Arbitrary<EditScriptStep> | null => {
  const indexes = blocks.flatMap((block, index) => (block.text.length >= 4 ? [index] : []));
  if (indexes.length === 0) {
    return null;
  }
  return fc.constantFrom(...indexes).chain((blockIndex) =>
    fc.record({
      type: fc.constant("formatRange" as const),
      blockIndex: fc.constant(blockIndex),
      startOffset: fc.constant(0),
      endOffset: fc.nat({ max: (blocks[blockIndex]?.text.length ?? 1) - 1 }).map((n) => n + 1),
      formatting: formattingArb,
    }),
  );
};

const editStepArb = (blocks: readonly FolioAIBlock[]): fc.Arbitrary<EditScriptStep> => {
  const steps: fc.Arbitrary<EditScriptStep>[] = [
    fc.record({
      type: fc.constant("insertParagraphAfter" as const),
      blockIndex: blockIndexArb(blocks),
      text: sentenceArb,
    }),
    fc.record({
      type: fc.constant("deleteParagraph" as const),
      blockIndex: blockIndexArb(blocks),
    }),
  ];

  const bodyIndex = bodyBlockIndexArb(blocks);
  if (bodyIndex) {
    steps.push(
      fc.record({
        type: fc.constant("moveParagraph" as const),
        blockIndex: bodyIndex,
        beforeBlockIndex: bodyIndex,
      }),
    );
  }

  const findable = findableWordArb(blocks);
  if (findable) {
    steps.push(
      findable.chain(({ blockIndex, find }) =>
        fc.record({
          type: fc.constant("replaceWords" as const),
          blockIndex: fc.constant(blockIndex),
          find: fc.constant(find),
          replace: wordArb,
        }),
      ),
    );
  }

  const formatStep = formatStepArb(blocks);
  if (formatStep) {
    steps.push(formatStep);
  }

  const tableIndex = tableBlockIndexArb(blocks);
  if (tableIndex) {
    steps.push(
      tableIndex.chain((blockIndex) =>
        fc.record({
          type: fc.constant("insertTableRow" as const),
          blockIndex: fc.constant(blockIndex),
          cellTexts: fc.array(wordArb, {
            minLength: rowBlockCount(blocks, blockIndex),
            maxLength: rowBlockCount(blocks, blockIndex),
          }),
        }),
      ),
      fc.record({ type: fc.constant("deleteTableRow" as const), blockIndex: tableIndex }),
      fc.record({
        type: fc.constant("editTableCell" as const),
        blockIndex: tableIndex,
        text: sentenceArb,
      }),
    );
  }

  return fc.oneof(...steps);
};

/**
 * Scripts whose steps address distinct blocks. Two edits to one block would
 * coalesce into a single reported change, which property 4 counts against the
 * script; keeping targets distinct lets it compare like with like.
 */
const editScriptArb = (blocks: readonly FolioAIBlock[]): fc.Arbitrary<EditScript> =>
  fc
    .array(editStepArb(blocks), { minLength: 1, maxLength: 4 })
    .map((steps) => distinctByBlock(steps));

const distinctByBlock = (steps: readonly EditScriptStep[]): EditScriptStep[] => {
  const seen = new Set<number>();
  const distinct: EditScriptStep[] = [];
  for (const step of steps) {
    const touched =
      step.type === "moveParagraph" ? [step.blockIndex, step.beforeBlockIndex] : [step.blockIndex];
    if (touched.some((index) => seen.has(index))) {
      continue;
    }
    for (const index of touched) {
      seen.add(index);
    }
    distinct.push(step);
  }
  return distinct;
};

/** The body-level blocks a story ENDS with, as a contiguous run of indexes. */
const trailingBodyIndexes = (blocks: readonly FolioAIBlock[]): number[] => {
  const indexes: number[] = [];
  for (let index = blocks.length - 1; index >= 0 && blocks[index]?.table === undefined; index--) {
    indexes.unshift(index);
  }
  return indexes;
};

/**
 * Scripts that remove the paragraphs a container ENDS with, with and without
 * new paragraphs written where they were.
 *
 * That shape is where the container's final paragraph mark cannot say what
 * happened: it is deleted nowhere, so the removal has to resolve onto the
 * paragraph it ends and anything put in its place has to land INSIDE it. The
 * general step arbitrary reaches the shape only by accident — it would have to
 * draw the last body block and every block back to the surviving one, and
 * never draw one twice — so it is generated on purpose here.
 *
 * The anchor covers both sides of the run: `"first"` puts the new paragraphs
 * where the removed ones started, `"last"` after the paragraph the container
 * ends with.
 */
const trailingRewriteScriptArb = (
  blocks: readonly FolioAIBlock[],
): fc.Arbitrary<EditScript> | null => {
  const tail = trailingBodyIndexes(blocks);
  if (tail.length === 0) {
    return null;
  }
  const lastIndex = tail.at(-1) ?? 0;
  return fc
    .record({
      // Zero removed is a plain append past the container's end, where the
      // mark that was ADDED has the same nowhere to go.
      removed: fc.integer({ min: 0, max: Math.min(tail.length, 4) }),
      inserted: fc.array(sentenceArb, { minLength: 0, maxLength: 3 }),
      anchor: fc.constantFrom("first" as const, "last" as const),
    })
    .filter(({ removed, inserted }) => removed > 0 || inserted.length > 0)
    .map(({ removed, inserted, anchor }) => {
      const run = tail.slice(tail.length - removed);
      const anchorIndex = (anchor === "first" ? run[0] : run.at(-1)) ?? lastIndex;
      const steps: EditScriptStep[] = inserted.map((text) => ({
        type: "insertParagraphAfter" as const,
        blockIndex: anchorIndex,
        text,
      }));
      for (const blockIndex of run) {
        steps.push({ type: "deleteParagraph", blockIndex });
      }
      return steps;
    });
};

/**
 * Every script shape the round trip has to hold for: the general edits, and
 * the trailing rewrites the general ones do not reach.
 */
const roundTripScriptArb = (blocks: readonly FolioAIBlock[]): fc.Arbitrary<EditScript> => {
  const trailing = trailingRewriteScriptArb(blocks);
  return trailing ? fc.oneof(editScriptArb(blocks), trailing) : editScriptArb(blocks);
};

const kindsOf = (changes: readonly CompareChange[]): string[] => changes.map(({ kind }) => kind);

/**
 * A block's text with the inline formatting it carries, one entry per
 * character, so two blocks compare on what a reader sees rather than on how
 * their runs happen to be split. A block the snapshot gives no runs for is
 * unstyled throughout.
 */
const inlineFormattingSignature = (block: FolioAIBlock): string =>
  JSON.stringify([
    block.text,
    (block.previewRuns ?? [{ text: block.text }]).flatMap((run) => {
      const formatting = [
        run.bold,
        run.italic,
        run.underline,
        run.strike,
        run.fontFamily,
        run.fontSizePt,
        run.color,
      ];
      return Array.from(run.text, () => formatting);
    }),
  ]);

const paragraphFormattingSignature = (block: FolioAIBlock): string =>
  JSON.stringify([
    block.text,
    block.styleId,
    block.listLevel,
    block.directAlignment,
    block.directSpacing,
  ]);

/** Every block a story holds, as signatures in an order-independent form. */
const signaturesOf = (
  blocks: readonly FolioAIBlock[],
  signature: (block: FolioAIBlock) => string,
): string => blocks.map(signature).toSorted().join("\n");

/**
 * How many independently reported formatting categories a relocation dropped.
 *
 * `moveParagraph` is a deletion plus an insertion of the block's TEXT: the
 * The edit-script vocabulary intentionally re-inserts only the paragraph's
 * text. Inline formatting can therefore produce one `format` change and direct
 * paragraph formatting can produce one `paragraph-format` change. Count those
 * projections separately so the change-count property stays tight as the
 * comparison learns to see another formatting category.
 *
 * Applied ON ITS OWN, and compared as an order-independent whole: a relocation
 * that carried its formatting leaves the very same paragraphs in a different
 * order, so any difference at all is formatting the re-insertion dropped. That
 * needs no guess about which arrival belongs to the step, which matching by
 * text alone cannot tell when a document repeats a paragraph.
 */
const relocationDroppedFormattingCount = async (
  base: ArrayBuffer,
  baseBlocks: readonly FolioAIBlock[],
  step: Extract<EditScriptStep, { type: "moveParagraph" }>,
): Promise<number> => {
  const relocated = await applyEditScript(base, [step]);
  if (relocated.isErr()) {
    throw relocated.error;
  }
  const relocatedBlocks = await blocksOf(relocated.value.buffer);
  const droppedInlineFormatting =
    signaturesOf(relocatedBlocks, inlineFormattingSignature) !==
    signaturesOf(baseBlocks, inlineFormattingSignature);
  const droppedParagraphFormatting =
    signaturesOf(relocatedBlocks, paragraphFormattingSignature) !==
    signaturesOf(baseBlocks, paragraphFormattingSignature);
  return Number(droppedInlineFormatting) + Number(droppedParagraphFormatting);
};

type TouchedBlockBudgetOptions = {
  base: ArrayBuffer;
  applied: readonly EditScriptStep[];
  baseBlocks: readonly FolioAIBlock[];
};

/**
 * Changes one script may legitimately produce.
 *
 * A paragraph step touches one block. A relocation touches two, and is
 * reported as two when the move pass declines to pair it (its text is below
 * the word floor, or it landed where the alignment can still walk forward).
 *
 * A relocation that also dropped formatting is budgeted one extra change for
 * each independently reported formatting projection. Every other script step,
 * including a whole row insertion or deletion, is one logical change. The
 * monotone row matcher keeps surviving cells paired instead of charging their
 * contents again when the table grows or shrinks.
 */
const touchedBlockBudget = async ({
  base,
  applied,
  baseBlocks,
}: TouchedBlockBudgetOptions): Promise<number> => {
  let budget = 0;
  for (const step of applied) {
    if (step.type !== "moveParagraph") {
      budget += 1;
      continue;
    }
    // oxlint-disable-next-line no-await-in-loop -- each relocation is replayed on its own
    budget += 2 + (await relocationDroppedFormattingCount(base, baseBlocks, step));
  }
  return budget;
};

describe("compareDocx", () => {
  test("base documents are present", () => {
    expect(BASE_DOCUMENTS.length).toBeGreaterThan(1);
  });

  test.each(BASE_CASES)(
    "comparing a document with itself reports nothing and changes nothing ($name)",
    async ({ buffer: base }) => {
      const { changes, buffer } = await compareOrThrow(base, base);
      expect(changes).toEqual([]);
      const [original, final, expected] = await Promise.all([
        projectView(buffer, "original"),
        projectView(buffer, "final"),
        projectView(base, "final"),
      ]);
      expect(original).toEqual(expected);
      expect(final).toEqual(expected);
    },
  );

  for (const { name, buffer: base, blocks: baseBlocks } of BASE_CASES) {
    test(
      `accepting every change yields the target and rejecting yields the base (${name})`,
      async () => {
        await fc.assert(
          fc.asyncProperty(roundTripScriptArb(baseBlocks), async (script) => {
            const scripted = await applyEditScript(base, script);
            if (scripted.isErr()) {
              throw scripted.error;
            }
            const target = scripted.value.buffer;
            const { buffer } = await compareOrThrow(base, target);
            const [accepted, rejected, targetProjection, baseProjection] = await Promise.all([
              projectView(buffer, "final"),
              projectView(buffer, "original"),
              projectView(target, "final"),
              projectView(base, "final"),
            ]);
            expect(accepted).toEqual(targetProjection);
            expect(rejected).toEqual(baseProjection);
          }),
          propertyConfig({ numRuns: 12 }),
        );
      },
      propertyTestTimeout(120_000),
    );

    test(
      `two runs produce byte-identical buffers and equal change lists (${name})`,
      async () => {
        await fc.assert(
          fc.asyncProperty(editScriptArb(baseBlocks), async (script) => {
            const scripted = await applyEditScript(base, script);
            if (scripted.isErr()) {
              throw scripted.error;
            }
            const target = scripted.value.buffer;
            // Sequential on purpose. The contract is that the same inputs give
            // the same output, not that two comparisons may share one realm:
            // the editor plugins the reviewer mounts are process singletons.
            const first = await compareOrThrow(base, target);
            const second = await compareOrThrow(base, target);
            expect(Buffer.from(first.buffer).equals(Buffer.from(second.buffer))).toBe(true);
            expect(first.changes).toEqual(second.changes);
          }),
          propertyConfig({ numRuns: 8 }),
        );
      },
      propertyTestTimeout(120_000),
    );

    test(
      `every revision id in the package is claimed once (${name})`,
      async () => {
        await fc.assert(
          fc.asyncProperty(editScriptArb(baseBlocks), async (script) => {
            const scripted = await applyEditScript(base, script);
            if (scripted.isErr()) {
              throw scripted.error;
            }
            const { buffer } = await compareOrThrow(base, scripted.value.buffer);
            const ids = await revisionIdsInPackage(buffer);
            expect(new Set(ids).size).toBe(ids.length);
          }),
          propertyConfig({ numRuns: 10 }),
        );
      },
      propertyTestTimeout(120_000),
    );

    test(
      `every paragraph id in the package fits the 31-bit range (${name})`,
      async () => {
        await fc.assert(
          fc.asyncProperty(editScriptArb(baseBlocks), async (script) => {
            const scripted = await applyEditScript(base, script);
            if (scripted.isErr()) {
              throw scripted.error;
            }
            const { buffer } = await compareOrThrow(base, scripted.value.buffer);
            const ids = await paragraphIdsInPackage(buffer);
            expect(ids.every((id) => Number.parseInt(id, 16) < 0x8000_0000)).toBe(true);
          }),
          propertyConfig({ numRuns: 10 }),
        );
      },
      propertyTestTimeout(120_000),
    );

    test(
      `the change count never exceeds the blocks the script touched (${name})`,
      async () => {
        await fc.assert(
          fc.asyncProperty(editScriptArb(baseBlocks), async (script) => {
            const scripted = await applyEditScript(base, script);
            if (scripted.isErr()) {
              throw scripted.error;
            }
            const { changes } = await compareOrThrow(base, scripted.value.buffer);
            expect(changes.length).toBeLessThanOrEqual(
              await touchedBlockBudget({ base, applied: scripted.value.applied, baseBlocks }),
            );
          }),
          propertyConfig({ numRuns: 12 }),
        );
      },
      propertyTestTimeout(120_000),
    );

    const formatOnly = formatStepArb(baseBlocks);
    if (formatOnly) {
      test(
        `a formatting-only script produces only format changes (${name})`,
        async () => {
          await fc.assert(
            fc.asyncProperty(
              fc.array(formatOnly, { minLength: 1, maxLength: 3 }).map(distinctByBlock),
              async (script) => {
                const scripted = await applyEditScript(base, script);
                if (scripted.isErr()) {
                  throw scripted.error;
                }
                if (scripted.value.applied.length === 0) {
                  return;
                }
                const { changes } = await compareOrThrow(base, scripted.value.buffer);
                expect(new Set(kindsOf(changes))).toEqual(
                  changes.length === 0 ? new Set() : new Set(["format"]),
                );
              },
            ),
            propertyConfig({ numRuns: 10 }),
          );
        },
        propertyTestTimeout(120_000),
      );
    }
  }

  test("a relocation past one neighbour reports the intended moved block", async () => {
    // Stable identity resolves the otherwise symmetric one-neighbour swap, so
    // the change list reports the scripted deletion and relocation rather than
    // describing both ends of the surviving neighbour as unrelated edits.
    const base = readFixture("upstream-styled-content.docx");
    const baseBlocks = await blocksOf(base);
    const move = { type: "moveParagraph", blockIndex: 1, beforeBlockIndex: 4 } as const;
    const script: EditScript = [{ type: "deleteParagraph", blockIndex: 2 }, move];

    // The scenario reads the fixture by index, so pin the shape it relies on:
    // the relocated paragraph carries direct formatting, and it swaps past
    // exactly one surviving neighbour.
    expect(baseBlocks.map(({ text }) => text)).toEqual([
      "Normal text. Bold text. Italic text. Underlined text.",
      "Bold and italic text. Strikethrough text.",
      "Large text (18pt). Small text (8pt).",
      "Centered paragraph.",
      "Right-aligned paragraph.",
    ]);
    const scripted = await applyEditScript(base, script);
    if (scripted.isErr()) {
      throw scripted.error;
    }
    expect(scripted.value.unresolved).toEqual([]);

    const { changes } = await compareOrThrow(base, scripted.value.buffer);
    expect(kindsOf(changes).toSorted()).toEqual(["delete", "move"]);
    expect(changes.find(({ kind }) => kind === "delete")).toMatchObject({
      baseBlockId: baseBlocks[2]?.id,
    });
    expect(changes.find(({ kind }) => kind === "move")).toMatchObject({
      baseBlockId: baseBlocks[1]?.id,
      text: baseBlocks[1]?.text,
    });
  });

  test("a paragraph inserted on a cell anchor lands beside the table it grew", async () => {
    // The second counterexample the change-count property found. An insertion
    // anchored in a cell writes its paragraph beside the table, so a script that
    // also changes that table's row count produces body changes the table's own
    // budget does not cover.
    const base = readFixture("upstream-with-tables.docx");
    const baseBlocks = await blocksOf(base);
    const script: EditScript = [
      { type: "insertParagraphAfter", blockIndex: 3, text: "Anchored in the third cell." },
      { type: "insertParagraphAfter", blockIndex: 4, text: "Anchored in the fourth cell." },
      { type: "insertTableRow", blockIndex: 1, cellTexts: ["one", "two", "three"] },
      { type: "insertTableRow", blockIndex: 2, cellTexts: ["four", "five", "six"] },
    ];

    // The scenario reads the fixture by index, so pin the shape it relies on:
    // blocks 1 to 4 are cells of one table with three columns.
    expect(baseBlocks.slice(1, 5).map(({ table }) => table?.tableIndex)).toEqual([0, 0, 0, 0]);

    const scripted = await applyEditScript(base, script);
    if (scripted.isErr()) {
      throw scripted.error;
    }
    expect(scripted.value.unresolved).toEqual([]);

    const targetBlocks = await blocksOf(scripted.value.buffer);
    expect(targetBlocks.filter(({ table }) => table === undefined).map(({ text }) => text)).toEqual(
      [
        "Document with tables:",
        "Anchored in the third cell.",
        "Anchored in the fourth cell.",
        "End of document.",
      ],
    );

    const { changes } = await compareOrThrow(base, scripted.value.buffer);
    expect(kindsOf(changes).filter((kind) => kind === "insert")).toHaveLength(2);
    expect(changes.length).toBeLessThanOrEqual(
      await touchedBlockBudget({ base, applied: scripted.value.applied, baseBlocks }),
    );
  });

  test("added rows do not make surviving cells report as changed", async () => {
    // Each inserted row is one structural change. Exact row evidence keeps the
    // pre-existing rows paired even when several insertions shift their indexes.
    const base = readFixture("upstream-with-tables.docx");
    const script: EditScript = [
      { type: "insertTableRow", blockIndex: 1, cellTexts: ["one", "two", "three"] },
      { type: "insertTableRow", blockIndex: 3, cellTexts: ["four", "five", "six"] },
      { type: "insertTableRow", blockIndex: 7, cellTexts: ["seven", "eight", "nine"] },
      { type: "insertTableRow", blockIndex: 9, cellTexts: ["ten", "eleven", "twelve"] },
    ];

    const scripted = await applyEditScript(base, script);
    if (scripted.isErr()) {
      throw scripted.error;
    }
    expect(scripted.value.unresolved).toEqual([]);

    const { changes } = await compareOrThrow(base, scripted.value.buffer);
    expect(kindsOf(changes)).toEqual([
      "table-row-insert",
      "table-row-insert",
      "table-row-insert",
      "table-row-insert",
    ]);
    expect(changes).toHaveLength(scripted.value.applied.length);
  });

  test.each([0, 1, 2] as const)(
    "deleting row %i keeps fully rewritten surviving rows at cell level",
    async (deletedRow) => {
      const base = readFixture("upstream-with-tables.docx");
      await fc.assert(
        fc.asyncProperty(rowRewriteArb, async (rewrittenText) => {
          const survivingRows = [0, 1, 2].filter((rowIndex) => rowIndex !== deletedRow);
          const script: EditScript = [
            ...survivingRows.flatMap((rowIndex) =>
              [0, 1, 2].map(
                (columnIndex): EditScriptStep => ({
                  type: "editTableCell",
                  blockIndex: 1 + rowIndex * 3 + columnIndex,
                  text: rewrittenText,
                }),
              ),
            ),
            { type: "deleteTableRow", blockIndex: 1 + deletedRow * 3 },
          ];

          const scripted = await applyEditScript(base, script);
          if (scripted.isErr()) {
            throw scripted.error;
          }
          expect(scripted.value.unresolved).toEqual([]);

          const { changes } = await compareOrThrow(base, scripted.value.buffer);
          expect(kindsOf(changes).toSorted()).toEqual([
            "replace",
            "replace",
            "replace",
            "replace",
            "replace",
            "replace",
            "table-row-delete",
          ]);
          expect(changes).toHaveLength(scripted.value.applied.length);
        }),
        propertyConfig({ numRuns: 4 }),
      );
    },
    propertyTestTimeout(120_000),
  );

  test.each([0, 1, 2, 3] as const)(
    "inserting row at position %i keeps fully rewritten existing rows at cell level",
    async (insertedRow) => {
      const base = readFixture("upstream-with-tables.docx");
      await fc.assert(
        fc.asyncProperty(rowRewriteArb, async (rewrittenText) => {
          const script: EditScript = [
            ...Array.from({ length: 9 }, (_unused, blockOffset) => ({
              type: "editTableCell" as const,
              blockIndex: 1 + blockOffset,
              text: rewrittenText,
            })),
            {
              type: "insertTableRow",
              blockIndex: insertedRow === 0 ? 1 : 1 + (insertedRow - 1) * 3,
              position:
                insertedRow === 0
                  ? EDIT_SCRIPT_TABLE_ROW_POSITION.before
                  : EDIT_SCRIPT_TABLE_ROW_POSITION.after,
              cellTexts: [rewrittenText, rewrittenText, rewrittenText],
            },
          ];

          const scripted = await applyEditScript(base, script);
          if (scripted.isErr()) {
            throw scripted.error;
          }
          expect(scripted.value.unresolved).toEqual([]);

          const { changes } = await compareOrThrow(base, scripted.value.buffer);
          const kinds = kindsOf(changes);
          expect(kinds.filter((kind) => kind === "replace")).toHaveLength(9);
          expect(kinds.filter((kind) => kind !== "replace")).toEqual(["table-row-insert"]);
          expect(changes).toHaveLength(scripted.value.applied.length);
        }),
        propertyConfig({ numRuns: 4 }),
      );
    },
    propertyTestTimeout(120_000),
  );

  for (const { name, buffer: base, blocks: baseBlocks } of BASE_CASES) {
    test(
      `no container's final paragraph mark carries any revision (${name})`,
      async () => {
        // A deleted paragraph mark means "merge this paragraph into the
        // following one" and an inserted one means the break was added, so
        // rejecting it closes the paragraph back over the next one. The last
        // paragraph of a body, a cell, a header, a note or a text box has no
        // following one, so neither states an edit a reader can carry out: one
        // makes a consumer refuse the package, the other leaves a revision
        // standing that neither accepting nor rejecting everything can clear.
        // Read back from the bytes the comparison produced, so it covers what
        // was written and not only what was planned.
        await fc.assert(
          fc.asyncProperty(roundTripScriptArb(baseBlocks), async (script) => {
            const scripted = await applyEditScript(base, script);
            if (scripted.isErr()) {
              throw scripted.error;
            }
            const { buffer } = await compareOrThrow(base, scripted.value.buffer);
            const written = await FolioDocxReviewer.fromBuffer(buffer);
            expect(revisedFinalParagraphMarks(written.toDocument())).toEqual([]);
          }),
          propertyConfig({ numRuns: 12 }),
        );
      },
      propertyTestTimeout(120_000),
    );
  }

  test(
    "a move-only script reports the relocation as a move",
    async () => {
      const base = SYNTHETIC_BASE;
      const baseBlocks = BASE_CASES.at(-1)?.blocks ?? [];
      const movable = movableBlockIndexArb(baseBlocks);
      if (movable === null) {
        throw new Error("The synthetic base offers no block the move pass would pair.");
      }
      await fc.assert(
        fc.asyncProperty(
          fc.tuple(movable, movable).filter(([from, to]) => Math.abs(from - to) > 1),
          async ([blockIndex, beforeBlockIndex]) => {
            const scripted = await applyEditScript(base, [
              { type: "moveParagraph", blockIndex, beforeBlockIndex },
            ]);
            if (scripted.isErr()) {
              throw scripted.error;
            }
            if (scripted.value.applied.length === 0) {
              return;
            }
            const { changes } = await compareOrThrow(base, scripted.value.buffer);
            // The alignment may absorb a move whose text still lands in an
            // order the LCS can walk forward; what it must never do is report
            // the relocation as unrelated churn.
            expect(kindsOf(changes).every((kind) => kind === "move")).toBe(true);
          },
        ),
        propertyConfig({ numRuns: 15 }),
      );
    },
    propertyTestTimeout(120_000),
  );

  describe("a base that already carries tracked changes", () => {
    test(
      "is compared as accepted, and the comparison is the only redline left",
      async () => {
        const revisedBase = await withPriorRevisions(SYNTHETIC_BASE);
        expect(await authorsOfChanges(revisedBase)).toEqual([PRIOR_AUTHOR]);
        const acceptedBase = await projectView(revisedBase, "final");

        await fc.assert(
          fc.asyncProperty(editScriptArb(await blocksOf(revisedBase)), async (script) => {
            const scripted = await applyEditScript(revisedBase, script);
            if (scripted.isErr()) {
              throw scripted.error;
            }
            const target = scripted.value.buffer;
            const { buffer } = await compareOrThrow(revisedBase, target);

            // Rejecting lands on the base as it stands, not on the document
            // before the previous reviewer touched it.
            expect(await projectView(buffer, "original")).toEqual(acceptedBase);
            expect(await projectView(buffer, "final")).toEqual(await projectView(target, "final"));

            // The prior reviewer's marks are resolved rather than layered
            // under this comparison's: a reader has one redline to read, and
            // rejecting it cannot land on a document neither side wrote.
            expect(await authorsOfChanges(buffer)).not.toContain(PRIOR_AUTHOR);
          }),
          propertyConfig({ numRuns: 10 }),
        );
      },
      propertyTestTimeout(120_000),
    );

    test("comparing it with itself still reports nothing", async () => {
      const revisedBase = await withPriorRevisions(SYNTHETIC_BASE);
      const { changes } = await compareOrThrow(revisedBase, revisedBase);
      expect(changes).toEqual([]);
    });
  });

  for (const { name, buffer: base, blocks: baseBlocks } of BASE_CASES) {
    test(
      `every scripted difference is representable, so the result verifies (${name})`,
      async () => {
        // The strict default turns an unproven redline into an error, so the
        // properties above enforce this by not throwing. Stated directly it
        // says the stronger thing: asked to emit whatever it can, the engine
        // still has nothing it could not represent. A scripted edit is built
        // from the operation vocabulary, so anything unverified here is a
        // defect in the engine rather than a document it cannot express.
        await fc.assert(
          fc.asyncProperty(roundTripScriptArb(baseBlocks), async (script) => {
            const scripted = await applyEditScript(base, script);
            if (scripted.isErr()) {
              throw scripted.error;
            }
            const result = await compareDocx(base, scripted.value.buffer, {
              ...OPTIONS,
              mode: "bestEffort",
            });
            if (result.isErr()) {
              throw result.error;
            }
            expect(result.value.verification).toEqual({ status: "verified" });
          }),
          propertyConfig({ numRuns: 10 }),
        );
      },
      propertyTestTimeout(120_000),
    );
  }
});
