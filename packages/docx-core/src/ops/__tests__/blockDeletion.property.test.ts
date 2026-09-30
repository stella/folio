/** Generated deletion laws cover content, survivor identity and terminal paragraph marks. */
import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { assertProperty, propertyTestTimeout } from "../../../../../test/property-testing";
import {
  paragraphNumberingReference,
  type BlockContent,
  type Document,
  type Paragraph,
} from "../../model/document";
import { applyDocumentOp, applyDocumentOps, type AppliedDocumentOp } from "../apply";
import { endsItsContainer, storyParagraphs } from "../blocks";
import { contractViolation } from "../contract";
import { revisionIdDemand } from "../plan";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import {
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
  type DeleteBlocksOp,
  type DocumentOp,
  type RevisionDecision,
} from "../types";

setDefaultTimeout(propertyTestTimeout(120_000));

const NUM_RUNS = 2_000;

const apply = (document: Document, op: DocumentOp) => {
  const result = applyDocumentOp(document, op);
  if (result.isErr()) throw result.error;
  return result.value;
};

const applyAll = (document: Document, ops: readonly DocumentOp[]) => {
  const result = applyDocumentOps(document, ops);
  if (result.isErr()) throw result.error;
  return result.value;
};

const resolve = (applied: AppliedDocumentOp, decision: RevisionDecision) =>
  apply(applied.document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: applied.revisions,
    decision,
  });

const expectInverse = (applied: AppliedDocumentOp, original: Document) => {
  const restored = applyAll(applied.document, applied.inverse);
  expect(restored.document).toStrictEqual(original);
  expect(applyAll(restored.document, restored.inverse).document).toStrictEqual(applied.document);
  expect(contractViolation(restored.document)).toBeUndefined();
};

const paragraphsOf = (document: Document) =>
  storyParagraphs(document.package.document).map(({ paragraph }) => paragraph);

const expectLocality = (original: Document, applied: AppliedDocumentOp) => {
  const touched = new Set([
    ...applied.touched.modified,
    ...applied.touched.inserted,
    ...applied.touched.removed,
  ]);
  const byId = new Map(
    paragraphsOf(applied.document).map((paragraph) => [paragraph.paraId, paragraph]),
  );
  for (const paragraph of paragraphsOf(original)) {
    if (!touched.has(paragraph.paraId ?? "")) expect(byId.get(paragraph.paraId)).toBe(paragraph);
  }
};

const expectFinalMarks = (document: Document) => {
  const body = document.package.document;
  expect(
    storyParagraphs(body).filter(
      (at) => endsItsContainer(body, at) && at.paragraph.pPrMark !== undefined,
    ),
  ).toEqual([]);
  expect(contractViolation(document)).toBeUndefined();
};

const paragraphArbitrary = fc
  .record({
    text: fc.string({ unit: fc.constantFrom("a", "b", " ", "ž", "😀"), maxLength: 12 }),
    alignment: fc.constantFrom("start", "end", "center"),
    bold: fc.boolean(),
    listed: fc.boolean(),
    markItalic: fc.boolean(),
    markHidden: fc.boolean(),
    runInWithNext: fc.boolean(),
  })
  .map(
    ({ text, alignment, bold, listed, markItalic, markHidden, runInWithNext }) =>
      ({
        type: "paragraph",
        formatting: {
          alignment,
          runProperties: { italic: markItalic, hidden: markHidden },
          runInWithNext,
          ...(listed ? { numPr: paragraphNumberingReference({ numId: 1, ilvl: 0 }) } : {}),
        },
        content:
          text === ""
            ? []
            : [{ type: "run", formatting: { bold }, content: [{ type: "text", text }] }],
      }) satisfies Paragraph,
  );

type Container = "body" | "cell" | "sdt" | "customXml";

const wrap = (content: BlockContent[], container: Container): BlockContent[] => {
  switch (container) {
    case "body":
      return content;
    case "cell":
      return [
        { type: "table", rows: [{ type: "tableRow", cells: [{ type: "tableCell", content }] }] },
      ];
    case "sdt":
      return [{ type: "blockSdt", properties: {}, content }];
    case "customXml":
      return [
        {
          type: "blockCustomXml",
          openingXml: '<w:customXml w:element="clause">',
          closingXml: "</w:customXml>",
          content,
        },
      ];
    default: {
      const unreachable: never = container;
      return unreachable;
    }
  }
};

const rangeArbitrary = fc.record({
  before: fc.array(paragraphArbitrary, { maxLength: 3 }),
  selected: fc.array(paragraphArbitrary, { minLength: 1, maxLength: 3 }),
  after: fc.array(paragraphArbitrary, { maxLength: 3 }),
});

type DeletionRange = ReturnType<typeof rangeArbitrary.generate>["value"];

const identifyRange = ({ before, selected, after }: DeletionRange, base: number) => {
  const originals = [...before, ...selected, ...after].map((paragraph, index) =>
    Object.assign({}, paragraph, { paraId: (base + index).toString(16).padStart(8, "0") }),
  );
  return {
    originals,
    selected: originals.slice(before.length, before.length + selected.length),
    before: originals.slice(0, before.length),
    after: originals.slice(before.length + selected.length),
  };
};

type IdentifiedRange = ReturnType<typeof identifyRange>;

const deletionOp = (range: IdentifiedRange, revisionId: number) =>
  ({
    type: DOCUMENT_OP_TYPES.DELETE_BLOCKS,
    story: OP_STORIES.MAIN,
    blockIds: range.selected.map(({ paraId }) => paraId),
    revision: { id: revisionId, author: "Reviewer", date: "2026-02-03T04:05:06Z" },
    newIds: { revision: Array.from({ length: 32 }, (_, index) => revisionId + index + 1) },
  }) satisfies DeleteBlocksOp;

const expectDirectRange = (range: IdentifiedRange, direct: AppliedDocumentOp) => {
  const originalIds = new Set(range.originals.map(({ paraId }) => paraId));
  const remaining = paragraphsOf(direct.document).filter(({ paraId }) =>
    originalIds.has(paraId ?? ""),
  );
  if (range.after.length > 0) {
    expect(remaining).toStrictEqual([...range.before, ...range.after]);
    for (const paragraph of [...range.before, ...range.after]) {
      expect(remaining.find(({ paraId }) => paraId === paragraph.paraId)).toBe(paragraph);
    }
    return;
  }
  const final = range.selected.at(-1);
  if (final === undefined) throw new Error("The generated selection is nonempty.");
  const preceding = range.before.at(-1);
  const survivor = remaining.at(-1);
  expect(survivor?.paraId).toBe(final.paraId);
  expect(remaining.slice(0, -1)).toStrictEqual(range.before.slice(0, -1));
  expect(survivor?.content).toStrictEqual(preceding?.content ?? []);
  const paragraphProps =
    preceding !== undefined && preceding.content.length > 0
      ? preceding.formatting
      : final.formatting;
  expect(survivor?.formatting).toStrictEqual({
    ...paragraphProps,
    runProperties: final.formatting.runProperties,
    runInWithNext: final.formatting.runInWithNext,
  });
};

const expectStaleInverse = (applied: AppliedDocumentOp) => {
  const replacement = applied.inverse.find((op) => op.type === DOCUMENT_OP_TYPES.REPLACE_BLOCKS);
  expect(replacement).toBeDefined();
  if (replacement === undefined || replacement.type !== DOCUMENT_OP_TYPES.REPLACE_BLOCKS) return;
  const expected = replacement.expected.at(0);
  if (expected?.paraId === undefined)
    throw new Error("The inverse must guard a surviving paragraph.");
  const changed = apply(applied.document, {
    type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
    story: OP_STORIES.MAIN,
    blockId: expected.paraId,
    patch: { styleId: "ChangedAfterDeletion" },
  });
  const restored = applyDocumentOps(changed.document, applied.inverse);
  expect(restored.isErr() ? restored.error.reason : undefined).toBe(
    DOCUMENT_OP_REFUSAL_REASONS.STALE,
  );
};

test("block deletion satisfies L1–L7, exact id demand and stale inverses", () => {
  assertProperty(
    fc.property(
      rangeArbitrary,
      fc.constantFrom("body", "cell", "sdt", "customXml"),
      (seed, container) => {
        const range = identifyRange(seed, 1);
        const outside: Paragraph = { type: "paragraph", paraId: "00000100", content: [] };
        const document: Document = {
          package: {
            document: {
              content:
                container === "body"
                  ? range.originals
                  : [...wrap(range.originals, container), outside],
            },
          },
        };
        const original = structuredClone(document);
        expectFinalMarks(document);
        const op = deletionOp(range, 1000);
        const demand = revisionIdDemand(document, op);
        if (demand.isErr()) throw demand.error;
        const exactOp = { ...op, newIds: { revision: op.newIds.revision.slice(0, demand.value) } };
        const tracked = apply(document, exactOp);
        const marked =
          range.after.length > 0
            ? range.selected
            : [...range.before.slice(-1), ...range.selected.slice(0, -1)];
        expect(
          paragraphsOf(tracked.document)
            .filter(({ pPrMark }) => pPrMark?.kind === "del")
            .map(({ paraId }) => paraId),
        ).toEqual(marked.map(({ paraId }) => paraId));
        const direct = apply(document, { ...op, revision: undefined });
        const accepted = resolve(tracked, REVISION_DECISIONS.ACCEPT);
        const rejected = resolve(tracked, REVISION_DECISIONS.REJECT);
        expect(accepted.document).toStrictEqual(direct.document); // L1
        expect(rejected.document).toStrictEqual(document); // L2
        expect(applyAll(document, [exactOp]).document).toStrictEqual(tracked.document); // L3
        for (const applied of [tracked, direct]) {
          expectInverse(applied, document); // L4
          expectLocality(document, applied); // L6
          expectFinalMarks(applied.document); // L7
          expectStaleInverse(applied);
        }
        expectInverse(accepted, tracked.document);
        expectInverse(rejected, tracked.document);
        expect(apply(structuredClone(document), structuredClone(exactOp))).toStrictEqual(tracked); // L5
        expectDirectRange(range, direct);
        expectFinalMarks(accepted.document);
        expectFinalMarks(rejected.document);
        for (const [decision, resolved] of [
          [REVISION_DECISIONS.ACCEPT, accepted],
          [REVISION_DECISIONS.REJECT, rejected],
        ] as const) {
          const twice = resolve({ ...tracked, document: resolved.document }, decision);
          expect(twice.document).toBe(resolved.document);
          expect(twice.inverse).toEqual([]);
        }
        expect(demand.value).toBe(Math.max(0, tracked.revisions.length - 1));
        if (demand.value > 0) {
          const short = applyDocumentOp(document, {
            ...op,
            newIds: { revision: op.newIds.revision.slice(0, demand.value - 1) },
          });
          expect(short.isErr() ? short.error.reason : undefined).toBe(
            DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS,
          );
        }
        expect(document).toStrictEqual(original);
      },
    ),
    { numRuns: NUM_RUNS },
  );
});

test("independent deletion sequences satisfy batch acceptance, rejection and inverse laws", () => {
  assertProperty(
    fc.property(
      fc.array(rangeArbitrary, { minLength: 2, maxLength: 3 }),
      fc.constantFrom("cell", "sdt", "customXml"),
      (seeds, container) => {
        const ranges = seeds.map((seed, index) => identifyRange(seed, 1 + index * 32));
        const outside: Paragraph = { type: "paragraph", paraId: "00000100", content: [] };
        const document: Document = {
          package: {
            document: {
              content: [...ranges.flatMap((range) => wrap(range.originals, container)), outside],
            },
          },
        };
        const ops = ranges.map((range, index) => deletionOp(range, 1000 + index * 100));
        const edits: AppliedDocumentOp[] = [];
        let current = document;
        let direct = document;
        for (const op of ops) {
          const tracked = apply(current, op);
          expect(resolve(tracked, REVISION_DECISIONS.REJECT).document).toStrictEqual(current);
          expect(resolve(tracked, REVISION_DECISIONS.ACCEPT).document).toStrictEqual(
            apply(current, { ...op, revision: undefined }).document,
          );
          expectInverse(tracked, current);
          expectLocality(current, tracked);
          expectFinalMarks(tracked.document);
          edits.push(tracked);
          current = tracked.document;
          direct = apply(direct, { ...op, revision: undefined }).document;
        }
        const batch = applyAll(document, ops);
        expect(batch.document).toStrictEqual(current);
        expectInverse(batch, document);
        const all = { ...batch, revisions: edits.flatMap(({ revisions }) => revisions) };
        expect(resolve(all, REVISION_DECISIONS.ACCEPT).document).toStrictEqual(direct);
        expect(resolve(all, REVISION_DECISIONS.REJECT).document).toStrictEqual(document);
        expect(
          applyAll(
            current,
            edits.toReversed().flatMap(({ inverse }) => inverse),
          ).document,
        ).toStrictEqual(document);
        let reversed = current;
        for (const edit of edits.toReversed()) {
          reversed = resolve({ ...edit, document: reversed }, REVISION_DECISIONS.REJECT).document;
        }
        expect(reversed).toStrictEqual(document);
      },
    ),
    { numRuns: NUM_RUNS },
  );
});
