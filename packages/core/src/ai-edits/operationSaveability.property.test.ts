/**
 * Class-level guard for issue #1103: no operation batch is ever answered as
 * applied while leaving a document the save refuses.
 *
 * Random batches over varied package shapes — no numbering part, bullets only,
 * decimals only, style-numbered headings, tables, comments, notes — mix valid
 * references with stale, foreign and malformed ones: block ids, numbering
 * instances, style ids, table anchors. Whatever the batch, mode, atomicity or
 * dry run, the invariant is the same: the batch is refused (at parse time, or
 * by a throw that leaves the document untouched), or its result reports every
 * skip as an issue — and either way the reviewer then saves and the saved
 * package reopens. "Committed, then unsaveable" is the one outcome that must
 * never happen.
 *
 * Deterministic: the seed is pinned, and the package shapes are built once.
 */

import { beforeAll, describe, expect, test } from "bun:test";
import fc from "fast-check";
import { readFileSync } from "node:fs";
import path from "node:path";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  FOLIO_DOCUMENT_OPERATION_MODES,
  InvalidFolioDocumentOperationBatchError,
  type FolioDocumentOperation,
  type FolioDocumentOperationMode,
} from "../document-operations";
import { fromMarkdown } from "../markdown/fromMarkdown";
import type { NumberingDefinitions } from "../types/document";
import { FolioDocxReviewer } from "./headless";
import type { FolioAIBlock } from "./types";

const FOOTNOTES_FIXTURE = path.join(
  import.meta.dir,
  "../docx/__tests__/__fixtures__/corpus/step3-footnotes.docx",
);

const decimalDefinition = (numId: number): NumberingDefinitions => ({
  abstractNums: [
    {
      abstractNumId: numId,
      multiLevelType: "multilevel",
      levels: [{ ilvl: 0, start: 1, numFmt: "decimal", lvlText: "%1.", suffix: "tab" }],
    },
  ],
  nums: [{ numId, abstractNumId: numId }],
});

const docxFromMarkdown = async (
  markdown: string,
  adjust?: (model: ReturnType<typeof fromMarkdown>) => void,
): Promise<Uint8Array> => {
  const model = fromMarkdown(markdown);
  adjust?.(model);
  return (await ensureParaIds(await createDocx(model))).docx;
};

const SHAPE_BUILDERS: Record<string, () => Promise<Uint8Array>> = {
  plain: () => docxFromMarkdown("# Title\n\nAlpha clause here.\n\nBeta clause here.\n\nGamma."),
  bullets: () => docxFromMarkdown("Intro text.\n\n- one item\n- two item\n- three item\n\nOutro."),
  decimals: () => docxFromMarkdown("Intro text.\n\n1. first item\n2. second item\n\nOutro."),
  styleNumberedHeadings: () =>
    docxFromMarkdown("# Heading one\n\nBody one.\n\n# Heading two\n\nBody two.", (model) => {
      model.package.numbering = decimalDefinition(7);
      const heading = model.package.styles?.styles.find(({ styleId }) => styleId === "Heading1");
      if (heading) {
        heading.pPr = { ...heading.pPr, numPr: { kind: "reference", numId: 7, ilvl: 0 } };
      }
    }),
  tables: () =>
    docxFromMarkdown(
      "Intro text.\n\n| Head A | Head B |\n| --- | --- |\n| cell c | cell d |\n| cell e | cell f |\n\nOutro.",
    ),
  mixed: () =>
    docxFromMarkdown(
      "Intro text.\n\n- bullet one\n- bullet two\n\n1. number one\n2. number two\n\n| A | B |\n| --- | --- |\n| c | d |\n\nOutro.",
    ),
  comments: async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(
      await docxFromMarkdown("Commented clause text.\n\nSecond clause.\n\nThird."),
      { author: "Reader" },
    );
    const target = reviewer.getContent()[0];
    reviewer.applyDocumentOperations({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "direct",
      operations: [
        { id: "c", type: "commentOnBlock", blockId: target?.id ?? "", comment: { text: "Note." } },
      ],
    });
    return new Uint8Array(await reviewer.toBuffer());
  },
  notes: async () => new Uint8Array(readFileSync(FOOTNOTES_FIXTURE)),
};

const SHAPES = Object.keys(SHAPE_BUILDERS);
const shapeBytes = new Map<string, Uint8Array>();
/** A block id read from a different document: valid-looking, not this one's. */
const foreignIds: string[] = [];

beforeAll(async () => {
  for (const shape of SHAPES) {
    const bytes = await SHAPE_BUILDERS[shape]!();
    shapeBytes.set(shape, bytes);
    const reviewer = await FolioDocxReviewer.fromBuffer(bytes, { author: "AI" });
    foreignIds.push(...reviewer.getContent().map((block) => block.id));
  }
});

type BlockRef =
  | { kind: "valid"; pick: number }
  | { kind: "tableCell"; pick: number }
  | { kind: "stale" }
  | { kind: "foreign"; pick: number }
  | { kind: "malformed" };

type NumberingRef =
  | { kind: "none" }
  | { kind: "clear" }
  | { kind: "defined"; pick: number; level: number }
  | { kind: "stale"; level: number }
  | { kind: "malformed" };

type StyleRef = { kind: "none" } | { kind: "clear" } | { kind: "named"; styleId: string };

const blockRefArb: fc.Arbitrary<BlockRef> = fc.oneof(
  { weight: 5, arbitrary: fc.record({ kind: fc.constant("valid" as const), pick: fc.nat(50) }) },
  {
    weight: 3,
    arbitrary: fc.record({ kind: fc.constant("tableCell" as const), pick: fc.nat(50) }),
  },
  { weight: 1, arbitrary: fc.constant({ kind: "stale" as const }) },
  { weight: 1, arbitrary: fc.record({ kind: fc.constant("foreign" as const), pick: fc.nat(200) }) },
  { weight: 1, arbitrary: fc.constant({ kind: "malformed" as const }) },
);

const numberingRefArb: fc.Arbitrary<NumberingRef> = fc.oneof(
  { weight: 3, arbitrary: fc.constant({ kind: "none" as const }) },
  { weight: 1, arbitrary: fc.constant({ kind: "clear" as const }) },
  {
    weight: 3,
    arbitrary: fc.record({
      kind: fc.constant("defined" as const),
      pick: fc.nat(10),
      level: fc.nat(9),
    }),
  },
  {
    weight: 3,
    arbitrary: fc.record({ kind: fc.constant("stale" as const), level: fc.nat(3) }),
  },
  { weight: 1, arbitrary: fc.constant({ kind: "malformed" as const }) },
);

const styleRefArb: fc.Arbitrary<StyleRef> = fc.oneof(
  { weight: 3, arbitrary: fc.constant({ kind: "none" as const }) },
  { weight: 1, arbitrary: fc.constant({ kind: "clear" as const }) },
  {
    weight: 2,
    arbitrary: fc.record({
      kind: fc.constant("named" as const),
      styleId: fc.constantFrom("Heading1", "Heading2", "ListParagraph", "NoSuchStyle", "Title"),
    }),
  },
);

const OPERATION_TYPES = [
  "insertAfterBlock",
  "insertBeforeBlock",
  "setBlockParagraphProperties",
  "replaceInBlock",
  "replaceBlock",
  "deleteBlock",
  "splitBlock",
  "mergeBlockWithNext",
  "commentOnBlock",
  "insertTable",
  "deleteTable",
  "insertSignatureTable",
  "insertTableRow",
  "deleteTableRow",
  "insertTableColumn",
  "deleteTableColumn",
  "mergeTableCells",
  "splitTableCell",
] as const;

type OperationSpec = {
  type: (typeof OPERATION_TYPES)[number];
  block: BlockRef;
  numbering: NumberingRef;
  style: StyleRef;
  listLevel: number | null | undefined;
  text: string;
};

const operationSpecArb: fc.Arbitrary<OperationSpec> = fc.record({
  type: fc.constantFrom(...OPERATION_TYPES),
  block: blockRefArb,
  numbering: numberingRefArb,
  style: styleRefArb,
  listLevel: fc.option(fc.option(fc.nat(3), { nil: null }), { nil: undefined }),
  text: fc.stringMatching(/^[A-Za-z][A-Za-z ]{0,14}$/u),
});

const batchArb = fc.record({
  shape: fc.constantFrom(...SHAPES),
  operations: fc.array(operationSpecArb, { minLength: 1, maxLength: 4 }),
  mode: fc.constantFrom<FolioDocumentOperationMode>(...FOLIO_DOCUMENT_OPERATION_MODES),
  atomic: fc.oneof({ weight: 3, arbitrary: fc.constant(false) }, fc.constant(true)),
  dryRun: fc.oneof({ weight: 4, arbitrary: fc.constant(false) }, fc.constant(true)),
});

type Materials = {
  blocks: FolioAIBlock[];
  cells: FolioAIBlock[];
  definedNumIds: number[];
};

const resolveBlockId = (ref: BlockRef, materials: Materials): string => {
  const pickFrom = (pool: readonly FolioAIBlock[], pick: number) =>
    pool.length === 0 ? "FFFFFFF0" : pool[pick % pool.length]!.id;
  switch (ref.kind) {
    case "valid":
      return pickFrom(materials.blocks, ref.pick);
    case "tableCell":
      return pickFrom(materials.cells.length > 0 ? materials.cells : materials.blocks, ref.pick);
    case "stale":
      return "7FFFFFF1";
    case "foreign":
      return foreignIds[ref.pick % foreignIds.length] ?? "7FFFFFF2";
    case "malformed":
      return "";
  }
};

const resolveNumbering = (
  ref: NumberingRef,
  materials: Materials,
): { numId: number; level: number } | null | undefined => {
  switch (ref.kind) {
    case "none":
      return undefined;
    case "clear":
      return null;
    case "defined": {
      const numId = materials.definedNumIds[ref.pick % Math.max(1, materials.definedNumIds.length)];
      return { numId: numId ?? 4242, level: ref.level };
    }
    case "stale":
      return { numId: Math.max(0, ...materials.definedNumIds) + 1000, level: ref.level };
    case "malformed":
      return { numId: 0, level: 0 };
  }
};

const resolveStyle = (ref: StyleRef): string | null | undefined => {
  if (ref.kind === "none") {
    return undefined;
  }
  return ref.kind === "clear" ? null : ref.styleId;
};

const materialize = (
  spec: OperationSpec,
  index: number,
  materials: Materials,
): FolioDocumentOperation => {
  const id = `op-${String(index)}`;
  const blockId = resolveBlockId(spec.block, materials);
  const numbering = resolveNumbering(spec.numbering, materials);
  const styleId = resolveStyle(spec.style);
  const block = materials.blocks.find((candidate) => candidate.id === blockId);
  const paragraph = {
    ...(styleId !== undefined && { styleId }),
    ...(numbering !== undefined && { numbering }),
    ...(spec.listLevel !== undefined && { listLevel: spec.listLevel }),
  };
  switch (spec.type) {
    case "insertAfterBlock":
    case "insertBeforeBlock":
      return { id, type: spec.type, blockId, text: spec.text, ...paragraph };
    case "setBlockParagraphProperties":
      return {
        id,
        type: spec.type,
        blockId,
        properties: Object.keys(paragraph).length > 0 ? paragraph : { alignment: "center" },
      };
    case "replaceInBlock": {
      const find = /[A-Za-z]{3,}/u.exec(block?.text ?? "")?.[0] ?? "missing";
      return { id, type: spec.type, blockId, find, replace: spec.text };
    }
    case "replaceBlock":
      return {
        id,
        type: spec.type,
        blockId,
        text: spec.text,
        ...(styleId !== undefined && { styleId }),
      };
    case "deleteBlock":
    case "deleteTable":
    case "deleteTableRow":
    case "deleteTableColumn":
    case "splitTableCell":
      return { id, type: spec.type, blockId };
    case "splitBlock":
      return {
        id,
        type: spec.type,
        blockId,
        offset: Math.min(2, Math.max(0, (block?.text.length ?? 1) - 1)),
        ...(Object.keys(paragraph).length > 0 && { secondParagraphProperties: paragraph }),
      };
    case "mergeBlockWithNext":
      return {
        id,
        type: spec.type,
        blockId,
        ...(Object.keys(paragraph).length > 0 && { mergedParagraphProperties: paragraph }),
      };
    case "commentOnBlock":
      return { id, type: spec.type, blockId, comment: { text: spec.text } };
    case "insertTable":
      return {
        id,
        type: spec.type,
        blockId,
        rows: [
          [spec.text, "b"],
          ["c", "d"],
        ],
      };
    case "insertSignatureTable":
      return { id, type: spec.type, blockId, parties: [{ name: spec.text }] };
    case "insertTableRow":
    case "insertTableColumn":
      return { id, type: spec.type, blockId, cellTexts: [spec.text] };
    case "mergeTableCells":
      return { id, type: spec.type, blockId, rowCount: 2 };
  }
};

const assertSaveable = async (reviewer: FolioDocxReviewer): Promise<void> => {
  const saved = await reviewer.toBuffer();
  const reopened = await FolioDocxReviewer.fromBuffer(saved, { author: "AI" });
  expect(reopened.getContent().length).toBeGreaterThan(0);
};

describe("operation batches never commit an unsaveable document", () => {
  test(
    "every batch is refused, or reports its skips, and the reviewer saves",
    async () => {
      await fc.assert(
        fc.asyncProperty(batchArb, async ({ shape, operations, mode, atomic, dryRun }) => {
          const reviewer = await FolioDocxReviewer.fromBuffer(shapeBytes.get(shape)!, {
            author: "AI",
          });
          const blocks = reviewer.getContent();
          const materials: Materials = {
            blocks,
            cells: blocks.filter((block) => block.table !== undefined),
            definedNumIds: [
              ...new Set(
                blocks.flatMap((block) =>
                  block.listReference === undefined ? [] : [block.listReference.numId],
                ),
              ),
            ],
          };
          const batch = {
            version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
            mode,
            atomic,
            dryRun,
            operations: operations.map((spec, index) => materialize(spec, index, materials)),
          };

          const before = reviewer.getContentAsText({ annotated: true });
          let result;
          try {
            result = reviewer.applyDocumentOperations(batch);
          } catch (error) {
            // A malformed reference is refused by the parser. Any other throw
            // (operations in one batch that contradict each other) is not this
            // property's subject, but it must still leave the document as it
            // was and saveable: a throw commits nothing.
            if (!(error instanceof InvalidFolioDocumentOperationBatchError)) {
              expect(reviewer.getContentAsText({ annotated: true })).toBe(before);
            }
            await assertSaveable(reviewer);
            return;
          }

          const issueCodes = result.issues.map(({ operationId, code }) => `${operationId}:${code}`);
          expect(issueCodes).toEqual(result.skipped.map(({ id, reason }) => `${id}:${reason}`));
          for (const { id } of result.applied) {
            expect(result.skipped.some((skip) => skip.id === id)).toBe(false);
          }
          await assertSaveable(reviewer);
        }),
        propertyConfig({ numRuns: 120, seed: 1103 }),
      );
    },
    propertyTestTimeout(120_000),
  );
});
