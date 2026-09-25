/**
 * The laws every schema-version-1 operation keeps, over synthetic documents
 * built from the model.
 *
 * 1. **Inverse.** Applying an operation and then its recorded inverse gives
 *    back a document structurally equal to the input, every unmodelled and
 *    captured field included, and so does the inverse after a JSON
 *    round-trip: it holds data, not references. Applying the inverse's own
 *    inverse redoes the operation. No inverse but a replacement's is a
 *    whole-paragraph replacement.
 * 2. **Sequences.** For a random sequence of operations, the inverses applied
 *    in reverse order restore the input exactly; a batch of the same
 *    operations gives the same document, and its inverse restores the input.
 * 3. **Determinism.** The same operation on equal documents gives equal
 *    results, including after the operation has been through JSON.
 * 4. **Locality.** A paragraph the operation does not report as touched is
 *    the same object afterwards, and so is every top-level block holding
 *    none, every package part besides the story, and the parser's section
 *    view of untouched blocks. The touched set names only ids the operation
 *    names.
 * 5. **Offsets.** Each operation changes the paragraph's logical text and run
 *    properties exactly as its definition says.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../../test/property-testing";
import type { BlockContent, Document, Paragraph, TextFormatting } from "../../model/document";
import { applyDocumentOp, applyDocumentOps, type AppliedDocumentOp } from "../apply";
import { storyParagraphs } from "../blocks";
import { sameRunFormatting } from "../inline";
import { childrenOf, isInlineContainer, isRemovedRevision, paragraphLogicalText } from "../offsets";
import { applyFormattingPatch } from "../patch";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type DocumentOpType,
  INHERIT_RUN_PROPS,
} from "../types";
import {
  documentArbitrary,
  GENERATED_OP_KINDS,
  type OpSeed,
  opFor,
  opSeedArbitrary,
} from "./documentArbitraries";

setDefaultTimeout(propertyTestTimeout(240_000));

const NUM_RUNS = 10_000;

type Tally = Map<DocumentOpType | "refused", number>;

const count = (tally: Tally, key: DocumentOpType | "refused"): void => {
  tally.set(key, (tally.get(key) ?? 0) + 1);
};

/**
 * A property that passes because nothing applied proves nothing: every
 * generated kind must have applied in a real share of the runs.
 */
const expectEveryKindApplied = (tally: Tally, runs: number): void => {
  for (const kind of GENERATED_OP_KINDS) {
    expect({ kind, applied: tally.get(kind) ?? 0 }).toEqual({
      kind,
      applied: expect.any(Number),
    });
    expect(tally.get(kind) ?? 0).toBeGreaterThan(runs / 1000);
  }
};

/** What each operation's inverse may be made of: the table in `apply.ts`. */
const INVERSE_KINDS = {
  insertText: ["deleteRange", "joinInline"],
  insertContent: ["deleteRange", "joinInline", "splitInline"],
  deleteRange: ["insertContent"],
  splitInline: ["joinInline"],
  joinInline: ["splitInline"],
  setRunProps: ["setRunProps", "joinInline"],
  setParagraphProps: ["setParagraphProps"],
  splitBlock: ["joinBlocks"],
  joinBlocks: ["splitBlock"],
  replaceBlocks: ["replaceBlocks"],
} as const satisfies Record<DocumentOpType, readonly DocumentOpType[]>;

const paragraphsById = (document: Document): Map<string, Paragraph> =>
  new Map(
    storyParagraphs(document.package.document).map(({ paragraph }) => [
      paragraph.paraId ?? "",
      paragraph,
    ]),
  );

const orderedIds = (document: Document): string[] =>
  storyParagraphs(document.package.document).map(({ paragraph }) => paragraph.paraId ?? "");

const blockHolds = (block: BlockContent, ids: ReadonlySet<string>): boolean => {
  switch (block.type) {
    case "paragraph":
      return block.paraId !== undefined && ids.has(block.paraId);
    case "table":
      return block.rows.some((row) =>
        row.cells.some((cell) => cell.content.some((child) => blockHolds(child, ids))),
      );
    case "blockSdt":
      return block.content.some((child) => blockHolds(child, ids));
    default:
      return false;
  }
};

const namedIds = (op: DocumentOp): Set<string> => {
  switch (op.type) {
    case DOCUMENT_OP_TYPES.INSERT_TEXT:
    case DOCUMENT_OP_TYPES.INSERT_CONTENT:
    case DOCUMENT_OP_TYPES.SPLIT_INLINE:
    case DOCUMENT_OP_TYPES.JOIN_INLINE:
      return new Set([op.at.blockId]);
    case DOCUMENT_OP_TYPES.DELETE_RANGE:
    case DOCUMENT_OP_TYPES.SET_RUN_PROPS:
      return new Set([op.from.blockId, op.to.blockId]);
    case DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS:
      return new Set([op.blockId]);
    case DOCUMENT_OP_TYPES.SPLIT_BLOCK:
      return new Set([op.at.blockId, op.newBlockId]);
    case DOCUMENT_OP_TYPES.JOIN_BLOCKS:
      return new Set([op.blockId, op.nextBlockId]);
    case DOCUMENT_OP_TYPES.REPLACE_BLOCKS:
      return new Set(
        [...op.expected, ...op.blocks].flatMap(({ paraId }) =>
          paraId === undefined ? [] : [paraId],
        ),
      );
    default: {
      const unreachable: never = op;
      return unreachable;
    }
  }
};

const touchedIds = ({ touched }: AppliedDocumentOp): Set<string> =>
  new Set([...touched.modified, ...touched.inserted, ...touched.removed]);

type Unit = { formatting: TextFormatting | undefined; inRun: boolean; removed: boolean };

/** The run properties and revision state of every unit, in offset order. */
const unitsOf = (paragraph: Paragraph): Unit[] => {
  const out: Unit[] = [];
  const walk = (items: Paragraph["content"], removed: boolean): void => {
    for (const item of items) {
      if (item.type === "run") {
        for (const content of item.content) {
          const width = content.type === "text" ? content.text.length : 1;
          for (let unit = 0; unit < width; unit += 1) {
            out.push({ formatting: item.formatting, inRun: true, removed });
          }
        }
        continue;
      }
      if (isInlineContainer(item)) {
        walk(childrenOf(item), removed || isRemovedRevision(item));
        continue;
      }
      if (paragraphLogicalText({ type: "paragraph", content: [item] }).length === 1) {
        out.push({ formatting: undefined, inRun: false, removed });
      }
    }
  };
  walk(paragraph.content, false);
  return out;
};

const applyAll = (document: Document, ops: readonly DocumentOp[]): Document => {
  const applied = applyDocumentOps(document, ops);
  if (applied.isErr()) {
    throw applied.error;
  }
  return applied.value.document;
};

const expectRestores = (applied: AppliedDocumentOp, original: Document): void => {
  const restored = applyDocumentOps(applied.document, applied.inverse);
  if (restored.isErr()) {
    throw restored.error;
  }
  expect(restored.value.document).toStrictEqual(original);
  // SAFETY: operations are plain data; this is the journal's round-trip.
  const replayed = JSON.parse(JSON.stringify(applied.inverse)) as DocumentOp[];
  expect(applyAll(applied.document, replayed)).toStrictEqual(original);
  expect(applyAll(restored.value.document, restored.value.inverse)).toStrictEqual(applied.document);
};

describe("document operations", () => {
  test("an operation's inverse restores the document exactly", () => {
    const tally: Tally = new Map();
    fc.assert(
      fc.property(documentArbitrary, opSeedArbitrary, (document, seed) => {
        const op = opFor(document, seed);
        const original = structuredClone(document);
        const applied = applyDocumentOp(document, op);
        // Applying, refused or not, never writes into its input.
        expect(document).toStrictEqual(original);
        if (applied.isErr()) {
          count(tally, "refused");
          return;
        }
        count(tally, op.type);
        const allowed: readonly DocumentOpType[] = INVERSE_KINDS[op.type];
        for (const inverse of applied.value.inverse) {
          expect(allowed).toContain(inverse.type);
        }
        expectRestores(applied.value, original);
      }),
      propertyConfig({ numRuns: NUM_RUNS }),
    );
    expectEveryKindApplied(tally, NUM_RUNS);
  });

  test("a sequence's inverses in reverse, and a batch's, restore the document exactly", () => {
    const tally: Tally = new Map();
    fc.assert(
      fc.property(
        documentArbitrary,
        fc.array(opSeedArbitrary, { minLength: 2, maxLength: 8 }),
        (document, seeds: OpSeed[]) => {
          const original = structuredClone(document);
          let current = document;
          const ops: DocumentOp[] = [];
          const inverses: (readonly DocumentOp[])[] = [];
          for (const seed of seeds) {
            const op = opFor(current, seed);
            const applied = applyDocumentOp(current, op);
            if (applied.isErr()) {
              count(tally, "refused");
              continue;
            }
            count(tally, op.type);
            ops.push(op);
            inverses.push(applied.value.inverse);
            current = applied.value.document;
          }
          expect(applyAll(current, inverses.toReversed().flat())).toStrictEqual(original);

          const batch = applyDocumentOps(document, ops);
          if (batch.isErr()) {
            throw batch.error;
          }
          expect(batch.value.document).toStrictEqual(current);
          expectRestores(batch.value, original);
        },
      ),
      propertyConfig({ numRuns: NUM_RUNS }),
    );
    expectEveryKindApplied(tally, NUM_RUNS);
  });

  test("the same operation on equal documents gives equal results", () => {
    const outcome = (result: ReturnType<typeof applyDocumentOp>) =>
      result.isOk() ? result.value : { refused: result.error.reason };
    fc.assert(
      fc.property(documentArbitrary, opSeedArbitrary, (document, seed) => {
        const op = opFor(document, seed);
        // SAFETY: the operation is plain data; this is the journal's round-trip.
        const replayed = JSON.parse(JSON.stringify(op)) as DocumentOp;
        const first = outcome(applyDocumentOp(structuredClone(document), op));
        expect(outcome(applyDocumentOp(structuredClone(document), replayed))).toStrictEqual(first);
        expect(outcome(applyDocumentOp(document, op))).toStrictEqual(first);
      }),
      propertyConfig({ numRuns: NUM_RUNS }),
    );
  });

  test("blocks an operation does not touch are the same objects", () => {
    const tally: Tally = new Map();
    fc.assert(
      fc.property(documentArbitrary, opSeedArbitrary, (document, seed) => {
        const op = opFor(document, seed);
        const applied = applyDocumentOp(document, op);
        if (applied.isErr()) {
          count(tally, "refused");
          return;
        }
        count(tally, op.type);
        const next = applied.value.document;
        const touched = touchedIds(applied.value);
        const named = namedIds(op);
        for (const id of touched) {
          expect(named.has(id)).toBe(true);
        }

        const before = paragraphsById(document);
        const after = paragraphsById(next);
        for (const [id, paragraph] of before) {
          if (!touched.has(id)) {
            expect(after.get(id)).toBe(paragraph);
          }
        }
        for (const [id, paragraph] of after) {
          if (!touched.has(id)) {
            expect(before.get(id)).toBe(paragraph);
          }
        }
        for (const id of applied.value.touched.modified) {
          expect(after.get(id)).not.toBe(before.get(id));
          expect(after.has(id)).toBe(true);
        }
        for (const id of applied.value.touched.inserted) {
          expect(before.has(id)).toBe(false);
          expect(after.has(id)).toBe(true);
        }
        for (const id of applied.value.touched.removed) {
          expect(after.has(id)).toBe(false);
        }
        const untouchedOrder = (ids: string[]) => ids.filter((id) => !touched.has(id));
        expect(untouchedOrder(orderedIds(next))).toEqual(untouchedOrder(orderedIds(document)));

        const untouchedBlocks = (target: Document) =>
          target.package.document.content.filter((block) => !blockHolds(block, touched));
        const beforeBlocks = untouchedBlocks(document);
        const afterBlocks = untouchedBlocks(next);
        expect(afterBlocks.length).toBe(beforeBlocks.length);
        for (const [index, block] of afterBlocks.entries()) {
          expect(block).toBe(beforeBlocks[index]!);
        }

        // The section view stays the body's blocks, and untouched sections stay put.
        const body = next.package.document;
        const beforeSections = document.package.document.sections ?? [];
        expect(body.sections?.flatMap(({ content }) => content)).toEqual(body.content);
        for (const [index, section] of (body.sections ?? []).entries()) {
          if (!section.content.some((block) => blockHolds(block, touched))) {
            expect(section).toBe(beforeSections[index]!);
          }
        }

        for (const key of Object.keys(document.package)) {
          if (key !== "document") {
            expect(Reflect.get(next.package, key)).toBe(Reflect.get(document.package, key));
          }
        }
        expect(body.finalSectionProperties).toBe(document.package.document.finalSectionProperties!);
        expect(body.comments).toBe(document.package.document.comments!);
        expect(next.warnings).toBe(document.warnings!);
      }),
      propertyConfig({ numRuns: NUM_RUNS }),
    );
    expectEveryKindApplied(tally, NUM_RUNS);
  });

  test("operations change logical text and run properties as defined", () => {
    fc.assert(
      fc.property(documentArbitrary, opSeedArbitrary, (document, seed) => {
        const op = opFor(document, seed);
        const applied = applyDocumentOp(document, op);
        if (applied.isErr()) {
          return;
        }
        const before = paragraphsById(document);
        const after = paragraphsById(applied.value.document);
        const paragraphOf = (paragraphs: Map<string, Paragraph>, id: string): Paragraph => {
          const paragraph = paragraphs.get(id);
          if (paragraph === undefined) throw new Error(`${id} is missing`);
          return paragraph;
        };
        const textOf = (paragraphs: Map<string, Paragraph>, id: string): string =>
          paragraphLogicalText(paragraphOf(paragraphs, id));
        switch (op.type) {
          case DOCUMENT_OP_TYPES.INSERT_TEXT: {
            const { blockId, offset } = op.at;
            const old = textOf(before, blockId);
            expect(textOf(after, blockId)).toBe(old.slice(0, offset) + op.text + old.slice(offset));
            const inserted = unitsOf(paragraphOf(after, blockId)).slice(
              offset,
              offset + op.text.length,
            );
            for (const unit of inserted) {
              expect(unit.inRun).toBe(true);
              expect(unit.removed).toBe(false);
              if (op.runProps !== INHERIT_RUN_PROPS) {
                expect(sameRunFormatting(unit.formatting, op.runProps)).toBe(true);
              }
            }
            break;
          }
          case DOCUMENT_OP_TYPES.INSERT_CONTENT: {
            const { blockId, offset } = op.at;
            const old = textOf(before, blockId);
            const added = paragraphLogicalText({
              type: "paragraph",
              content: [...op.slice.content],
            });
            expect(textOf(after, blockId)).toBe(old.slice(0, offset) + added + old.slice(offset));
            break;
          }
          case DOCUMENT_OP_TYPES.DELETE_RANGE: {
            const { blockId } = op.from;
            const old = textOf(before, blockId);
            expect(textOf(after, blockId)).toBe(
              old.slice(0, op.from.offset) + old.slice(op.to.offset),
            );
            break;
          }
          case DOCUMENT_OP_TYPES.SPLIT_INLINE:
          case DOCUMENT_OP_TYPES.JOIN_INLINE: {
            const { blockId } = op.at;
            expect(textOf(after, blockId)).toBe(textOf(before, blockId));
            expect(unitsOf(paragraphOf(after, blockId))).toEqual(
              unitsOf(paragraphOf(before, blockId)),
            );
            break;
          }
          case DOCUMENT_OP_TYPES.SET_RUN_PROPS: {
            const { blockId } = op.from;
            expect(textOf(after, blockId)).toBe(textOf(before, blockId));
            const oldUnits = unitsOf(paragraphOf(before, blockId));
            const newUnits = unitsOf(paragraphOf(after, blockId));
            for (const [index, unit] of newUnits.entries()) {
              const old = oldUnits[index]!;
              const inRange = index >= op.from.offset && index < op.to.offset;
              const expected =
                inRange && old.inRun
                  ? applyFormattingPatch(old.formatting, op.patch)
                  : old.formatting;
              expect(sameRunFormatting(unit.formatting, expected)).toBe(true);
            }
            break;
          }
          case DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS: {
            expect(after.get(op.blockId)?.content).toBe(paragraphOf(before, op.blockId).content);
            break;
          }
          case DOCUMENT_OP_TYPES.SPLIT_BLOCK: {
            const { blockId, offset } = op.at;
            const old = textOf(before, blockId);
            expect(textOf(after, blockId)).toBe(old.slice(0, offset));
            expect(textOf(after, op.newBlockId)).toBe(old.slice(offset));
            break;
          }
          case DOCUMENT_OP_TYPES.JOIN_BLOCKS: {
            expect(textOf(after, op.blockId)).toBe(
              textOf(before, op.blockId) + textOf(before, op.nextBlockId),
            );
            break;
          }
          case DOCUMENT_OP_TYPES.REPLACE_BLOCKS:
            break;
          default: {
            const unreachable: never = op;
            return unreachable;
          }
        }
      }),
      propertyConfig({ numRuns: NUM_RUNS }),
    );
  });
});
