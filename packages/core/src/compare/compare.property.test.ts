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
import { applyEditScript, type EditScript, type EditScriptStep } from "./scenario";
import type { CompareChange, CompareResult } from "./types";
import { deletedFinalParagraphMarks } from "./verification";

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
  for (const [path, file] of Object.entries(zip.files)) {
    if (file.dir || !path.startsWith("word/") || !path.endsWith(".xml")) {
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

const kindsOf = (changes: readonly CompareChange[]): string[] => changes.map(({ kind }) => kind);

/**
 * Changes one script may legitimately produce.
 *
 * A paragraph step touches one block. A relocation touches two, and is
 * reported as two when the move pass declines to pair it (its text is below
 * the word floor, or it landed where the alignment can still walk forward).
 *
 * A table whose row COUNT changed is budgeted at its whole size. Rows pair on
 * exact text first and positionally after that, so once the counts differ the
 * surviving rows can line up one row off and every cell in the table reports as
 * changed. The result still accepts back to the target — that is a separate
 * property — but it is more granular than the script was. Bounding it at one
 * table keeps the real guarantee: the comparison never invents work beyond the
 * content the script disturbed.
 */
const touchedBlockBudget = (
  applied: readonly EditScriptStep[],
  blocks: readonly FolioAIBlock[],
): number => {
  const rowCountChanged = new Set(
    applied.flatMap((step) => {
      if (step.type !== "insertTableRow" && step.type !== "deleteTableRow") {
        return [];
      }
      const tableIndex = blocks[step.blockIndex]?.table?.tableIndex;
      return tableIndex === undefined ? [] : [tableIndex];
    }),
  );
  let budget = 0;
  for (const tableIndex of rowCountChanged) {
    budget += blocks.filter((block) => block.table?.tableIndex === tableIndex).length;
  }
  for (const step of applied) {
    const tableIndex = blocks[step.blockIndex]?.table?.tableIndex;
    if (tableIndex !== undefined && rowCountChanged.has(tableIndex)) {
      continue;
    }
    budget += step.type === "moveParagraph" ? 2 : 1;
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
          fc.asyncProperty(editScriptArb(baseBlocks), async (script) => {
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
              touchedBlockBudget(scripted.value.applied, baseBlocks),
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

  for (const { name, buffer: base, blocks: baseBlocks } of BASE_CASES) {
    test(
      `no container's final paragraph mark is ever deleted (${name})`,
      async () => {
        // A deleted paragraph mark means "merge this paragraph into the
        // following one", and the last paragraph of a body, a cell, a header,
        // a note or a text box has no following one: a consumer reading such a
        // mark refuses the package rather than opening it. Read back from the
        // bytes the comparison produced, so it covers what was written and not
        // only what was planned.
        await fc.assert(
          fc.asyncProperty(editScriptArb(baseBlocks), async (script) => {
            const scripted = await applyEditScript(base, script);
            if (scripted.isErr()) {
              throw scripted.error;
            }
            const { buffer } = await compareOrThrow(base, scripted.value.buffer);
            const written = await FolioDocxReviewer.fromBuffer(buffer);
            expect(deletedFinalParagraphMarks(written.toDocument())).toEqual([]);
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
          fc.asyncProperty(editScriptArb(baseBlocks), async (script) => {
            const scripted = await applyEditScript(base, script);
            if (scripted.isErr()) {
              throw scripted.error;
            }
            const result = await compareDocx(base, scripted.value.buffer, {
              ...OPTIONS,
              onUnverified: "emit",
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
