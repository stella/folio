/**
 * The laws tracked operations keep, over review-valid synthetic documents
 * (marks of every kind but none on a paragraph that ends its container,
 * deletions nested in insertions). Each tracked operation carries a stamp
 * dated after every change in the document, so it records changes of its own
 * rather than merging into one that was there.
 *
 * - **L4 Inverse.** Tracked operations are undone exactly by their inverses,
 *   after a JSON round trip too.
 * - **L5 Determinism.** Equal inputs give equal results.
 * - **L6 Locality.** Paragraphs an operation does not touch are the same
 *   objects; it touches only paragraphs it names.
 * - **L7 Review shape.** No tracked operation leaves a mark on a paragraph
 *   that ends its container, and the seed contract holds.
 *
 * `revisionIdDemand` counts exactly the new ids an operation takes, and a
 * planned tracked deletion removes only the author's own insertions.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../../test/property-testing";
import type { BlockContent, Document, Paragraph } from "../../model/document";
import { applyDocumentOp, applyDocumentOps, type AppliedDocumentOp } from "../apply";
import { storyParagraphs } from "../blocks";
import { contractViolation } from "../contract";
import { idKey } from "../ids";
import { compareGaps, type Gap, leafSpans } from "../leaves";
import { paragraphLogicalText } from "../offsets";
import { planTrackedDeletion, revisionIdDemand } from "../plan";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import { isTrackedWrapper } from "../review";
import { DOCUMENT_OP_TYPES, type DocumentOp, type DocumentOpType } from "../types";
import {
  GENERATED_TRACKED_OP_KINDS,
  independentCopy,
  opSeedArbitrary,
  reviewDocumentArbitrary,
  trackedOpFor,
} from "./documentArbitraries";

setDefaultTimeout(propertyTestTimeout(240_000));

/**
 * Each run applies an operation and checks it several ways, so the base count
 * is kept small: the nightly factor of ten takes it to 10^4, and the five-fold
 * run for changed areas stays within that job's time budget.
 */
const NUM_RUNS = 1000;

type Tally = Map<string, number>;

const count = (tally: Tally, key: string): void => {
  tally.set(key, (tally.get(key) ?? 0) + 1);
};

/** A law that holds because nothing applied proves nothing. */
const expectEveryKindChecked = (tally: Tally, runs: number): void => {
  for (const kind of GENERATED_TRACKED_OP_KINDS) {
    expect({ kind, checked: (tally.get(kind) ?? 0) > runs / 2000 }).toEqual({
      kind,
      checked: true,
    });
  }
};

const kindOf = (op: DocumentOp): DocumentOpType => op.type;

const paragraphById = (document: Document, id: string): Paragraph | undefined =>
  storyParagraphs(document.package.document).find(
    ({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(id),
  )?.paragraph;

/** The leaves of a paragraph between two gaps. */
const leavesBetween = (paragraph: Paragraph, from: Gap, to: Gap) =>
  leafSpans(paragraph.content).filter((span) => {
    const start = compareGaps(span.before, from) >= 0 ? span.before : from;
    const end = compareGaps(span.after, to) <= 0 ? span.after : to;
    return compareGaps(start, end) < 0;
  });

/** Paragraphs carrying a mark that end their story body or table cell. */
const containerFinalMarks = (document: Document): string[] => {
  const out: string[] = [];
  const visit = (list: readonly BlockContent[], endsContainer: boolean): void => {
    for (const [index, block] of list.entries()) {
      const last = index === list.length - 1;
      switch (block.type) {
        case "paragraph":
          if (last && endsContainer && block.pPrMark !== undefined) out.push(block.paraId ?? "");
          break;
        case "table":
          for (const row of block.rows) for (const cell of row.cells) visit(cell.content, true);
          break;
        case "blockSdt":
        case "blockCustomXml":
          visit(block.content, last && endsContainer);
          break;
        default:
          break;
      }
    }
  };
  visit(document.package.document.content, true);
  return out;
};

const applyAll = (document: Document, ops: readonly DocumentOp[]): Document => {
  const applied = applyDocumentOps(document, ops);
  if (applied.isErr()) throw applied.error;
  return applied.value.document;
};

/** L4: the inverse restores exactly, after a JSON round trip too, and redoes. */
const expectRestores = (applied: AppliedDocumentOp, original: Document): void => {
  expect(contractViolation(applied.document)).toBeUndefined();
  const restored = applyDocumentOps(applied.document, applied.inverse);
  if (restored.isErr()) throw restored.error;
  expect(restored.value.document).toStrictEqual(original);
  // SAFETY: operations are plain data; this is the journal's round trip.
  const replayed = JSON.parse(JSON.stringify(applied.inverse)) as DocumentOp[];
  expect(applyAll(applied.document, replayed)).toStrictEqual(original);
  expect(applyAll(restored.value.document, restored.value.inverse)).toStrictEqual(applied.document);
};

const outcome = (result: ReturnType<typeof applyDocumentOp>) =>
  result.isOk() ? result.value : { refused: result.error.reason };

/** The paragraphs an operation names. */
const namedParagraphs = (op: DocumentOp): Set<string> => {
  switch (op.type) {
    case DOCUMENT_OP_TYPES.INSERT_TEXT:
    case DOCUMENT_OP_TYPES.INSERT_CONTENT:
    case DOCUMENT_OP_TYPES.SPLIT_INLINE:
    case DOCUMENT_OP_TYPES.JOIN_INLINE:
      return new Set([idKey(op.at.blockId)]);
    case DOCUMENT_OP_TYPES.DELETE_RANGE:
    case DOCUMENT_OP_TYPES.SET_RUN_PROPS:
      return new Set([idKey(op.from.blockId), idKey(op.to.blockId)]);
    case DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS:
    case DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW:
    case DOCUMENT_OP_TYPES.REPLACE_INLINE:
      return new Set([idKey(op.blockId)]);
    case DOCUMENT_OP_TYPES.SPLIT_BLOCK:
      return new Set([idKey(op.at.blockId), idKey(op.newBlockId)]);
    case DOCUMENT_OP_TYPES.JOIN_BLOCKS:
      return new Set([idKey(op.blockId), idKey(op.nextBlockId)]);
    case DOCUMENT_OP_TYPES.REPLACE_BLOCKS:
      return new Set(
        [...op.expected, ...op.blocks].flatMap(({ paraId }) =>
          paraId === undefined ? [] : [idKey(paraId)],
        ),
      );
    default: {
      const unreachable: never = op;
      return unreachable;
    }
  }
};

describe("tracked operations", () => {
  test("L4: a tracked operation is undone exactly by its inverse", () => {
    const tally: Tally = new Map();
    fc.assert(
      fc.property(reviewDocumentArbitrary, opSeedArbitrary, (document, seed) => {
        const original = structuredClone(document);
        const op = trackedOpFor(document, seed);
        const applied = applyDocumentOp(document, op);
        // Applying, refused or not, never writes into its input.
        expect(document).toStrictEqual(original);
        if (applied.isErr()) {
          count(tally, "refused");
          return;
        }
        count(tally, kindOf(op));
        expectRestores(applied.value, original);
      }),
      propertyConfig({ numRuns: NUM_RUNS }),
    );
    expectEveryKindChecked(tally, NUM_RUNS);
  });

  test("L4: a run of tracked operations is undone exactly, in reverse and as a batch", () => {
    const tally: Tally = new Map();
    fc.assert(
      fc.property(
        reviewDocumentArbitrary,
        fc.array(opSeedArbitrary, { minLength: 2, maxLength: 6 }),
        (document, seeds) => {
          const original = structuredClone(document);
          let current = document;
          const ops: DocumentOp[] = [];
          const inverses: (readonly DocumentOp[])[] = [];
          for (const [index, seed] of seeds.entries()) {
            const op = trackedOpFor(current, seed, index);
            const applied = applyDocumentOp(current, op);
            if (applied.isErr()) continue;
            count(tally, kindOf(op));
            ops.push(op);
            inverses.push(applied.value.inverse);
            current = applied.value.document;
          }
          expect(applyAll(current, inverses.toReversed().flat())).toStrictEqual(original);
          const batch = applyDocumentOps(document, ops);
          if (batch.isErr()) throw batch.error;
          expect(batch.value.document).toStrictEqual(current);
          expectRestores(batch.value, original);
        },
      ),
      propertyConfig({ numRuns: NUM_RUNS / 5 }),
    );
    expectEveryKindChecked(tally, NUM_RUNS / 5);
  });

  test("L5: equal inputs give equal results", () => {
    fc.assert(
      fc.property(reviewDocumentArbitrary, opSeedArbitrary, (document, seed) => {
        const op = trackedOpFor(document, seed);
        // SAFETY: the operation is plain data; this is the journal's round trip.
        const replayed = JSON.parse(JSON.stringify(op)) as DocumentOp;
        const first = outcome(applyDocumentOp(structuredClone(document), op));
        expect(outcome(applyDocumentOp(document, replayed))).toStrictEqual(first);
        expect(outcome(applyDocumentOp(independentCopy(document), replayed))).toStrictEqual(first);
      }),
      propertyConfig({ numRuns: NUM_RUNS }),
    );
  });

  test("L6: a tracked operation touches only the paragraphs it names", () => {
    fc.assert(
      fc.property(reviewDocumentArbitrary, opSeedArbitrary, (document, seed) => {
        const op = trackedOpFor(document, seed);
        const applied = applyDocumentOp(document, op);
        if (applied.isErr()) return;
        const named = namedParagraphs(op);
        const touched = new Set(
          [
            ...applied.value.touched.modified,
            ...applied.value.touched.inserted,
            ...applied.value.touched.removed,
          ].map(idKey),
        );
        for (const id of touched) expect(named.has(id)).toBe(true);
        const after = new Map(
          storyParagraphs(applied.value.document.package.document).map(({ paragraph }) => [
            idKey(paragraph.paraId ?? ""),
            paragraph,
          ]),
        );
        for (const { paragraph } of storyParagraphs(document.package.document)) {
          const key = idKey(paragraph.paraId ?? "");
          if (!touched.has(key)) expect(after.get(key)).toBe(paragraph);
        }
      }),
      propertyConfig({ numRuns: NUM_RUNS }),
    );
  });

  test("L7: no tracked operation marks a paragraph that ends its container", () => {
    const tally: Tally = new Map();
    fc.assert(
      fc.property(reviewDocumentArbitrary, opSeedArbitrary, (document, seed) => {
        expect(containerFinalMarks(document)).toEqual([]);
        const op = trackedOpFor(document, seed);
        const applied = applyDocumentOp(document, op);
        if (applied.isErr()) return;
        count(tally, kindOf(op));
        expect(containerFinalMarks(applied.value.document)).toEqual([]);
        expect(contractViolation(applied.value.document)).toBeUndefined();
      }),
      propertyConfig({ numRuns: NUM_RUNS }),
    );
    expectEveryKindChecked(tally, NUM_RUNS);
  });

  test("a tracked operation takes exactly the new ids revisionIdDemand counts", () => {
    fc.assert(
      fc.property(reviewDocumentArbitrary, opSeedArbitrary, (document, seed) => {
        const op = trackedOpFor(document, { ...seed, depth: 1 });
        const demand = revisionIdDemand(document, op);
        if (demand.isErr()) return;
        const base = 900_000;
        const withIds = (length: number): DocumentOp => ({
          ...op,
          newIds: {
            revision: Array.from({ length }, (_, index) => base + index),
            control: Array.from({ length: 64 }, (_, index) => base + index),
          },
        });
        expect(applyDocumentOp(document, withIds(demand.value)).isOk()).toBe(true);
        if (demand.value > 0) {
          const short = applyDocumentOp(document, withIds(demand.value - 1));
          expect(short.isErr() && short.error.reason).toBe(
            DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS,
          );
        }
      }),
      propertyConfig({ numRuns: NUM_RUNS / 5 }),
    );
  });

  test("a planned tracked deletion removes only the author's own insertions", () => {
    const tally: Tally = new Map();
    fc.assert(
      fc.property(reviewDocumentArbitrary, opSeedArbitrary, (document, seed) => {
        const op = trackedOpFor(document, { ...seed, kind: 2 });
        if (op.type !== DOCUMENT_OP_TYPES.DELETE_RANGE || op.revision === undefined) return;
        const plan = planTrackedDeletion(document, {
          from: op.from,
          to: op.to,
          revision: op.revision,
          newIds: { revision: Array.from({ length: 64 }, (_, index) => 700_000 + index) },
        });
        if (plan.isErr()) {
          count(tally, "refused");
          return;
        }
        const applied = applyDocumentOps(document, plan.value);
        if (applied.isErr()) throw applied.error;
        count(tally, plan.value.length > 1 ? "several" : "one");
        const paragraph = paragraphById(document, op.from.blockId);
        if (paragraph === undefined) return;
        const text = paragraphLogicalText(paragraph);
        const from = {
          offset: op.from.offset,
          zeroWidthBefore: op.from.zeroWidthBefore ?? Number.MAX_SAFE_INTEGER,
        };
        const to = { offset: op.to.offset, zeroWidthBefore: op.to.zeroWidthBefore ?? 0 };
        const own = new Set<number>();
        for (const span of leavesBetween(paragraph, from, to)) {
          const removed = span.ancestors.some(
            (ancestor) => ancestor.type === "deletion" || ancestor.type === "moveFrom",
          );
          const mine = span.ancestors.some(
            (ancestor) =>
              ancestor.type === "insertion" &&
              isTrackedWrapper(ancestor) &&
              ancestor.info.author === op.revision?.author,
          );
          if (removed || !mine) continue;
          const start = Math.max(span.before.offset, op.from.offset);
          const end = Math.min(span.after.offset, op.to.offset);
          for (let unit = start; unit < end; unit += 1) own.add(unit);
        }
        const expected = Array.from({ length: text.length }, (_, unit) =>
          own.has(unit) ? "" : text.charAt(unit),
        ).join("");
        const after = paragraphById(applied.value.document, op.from.blockId);
        expect(after === undefined ? undefined : paragraphLogicalText(after)).toBe(expected);
        expectRestores(applied.value, structuredClone(document));
      }),
      propertyConfig({ numRuns: NUM_RUNS / 5 }),
    );
    expect(tally.get("one") ?? 0).toBeGreaterThan(0);
  });
});
