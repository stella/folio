/** Block insertion laws, including exact restoration of the untouched boundary paragraph. */
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
import {
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
  type DocumentOp,
  type InsertBlocksOp,
  type RevisionDecision,
} from "../types";

setDefaultTimeout(propertyTestTimeout(120_000));

const apply = (document: Document, op: DocumentOp) => {
  const applied = applyDocumentOp(document, op);
  if (applied.isErr()) throw applied.error;
  return applied.value;
};

const resolve = (applied: AppliedDocumentOp, decision: RevisionDecision) =>
  apply(applied.document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: applied.revisions,
    decision,
  });

const undo = (applied: AppliedDocumentOp) => {
  const restored = applyDocumentOps(applied.document, applied.inverse);
  if (restored.isErr()) throw restored.error;
  return restored.value.document;
};

const paragraphArbitrary = fc
  .record({
    text: fc.string({ unit: fc.constantFrom("a", "b", " ", "ž", "😀"), maxLength: 16 }),
    alignment: fc.constantFrom("start", "end", "center"),
    bold: fc.boolean(),
    listed: fc.boolean(),
  })
  .map(
    ({ text, alignment, bold, listed }) =>
      ({
        type: "paragraph",
        formatting: {
          alignment,
          ...(listed ? { numPr: paragraphNumberingReference({ numId: 1, ilvl: 0 }) } : {}),
        },
        content:
          text === ""
            ? []
            : [{ type: "run", formatting: { bold }, content: [{ type: "text", text }] }],
      }) satisfies Paragraph,
  );

const wrap = (
  content: BlockContent[],
  container: "body" | "cell" | "sdt" | "customXml",
): BlockContent[] => {
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

const caseArbitrary = fc.record({
  originals: fc.array(paragraphArbitrary, { minLength: 2, maxLength: 4 }),
  batches: fc.array(fc.array(paragraphArbitrary, { minLength: 1, maxLength: 3 }), {
    minLength: 1,
    maxLength: 3,
  }),
  container: fc.constantFrom("body", "cell", "sdt", "customXml"),
  after: fc.boolean(),
});

test("block insertion satisfies L1–L7 and idempotent review", () => {
  assertProperty(
    fc.property(caseArbitrary, ({ originals, batches, container, after }) => {
      const paragraphs = originals.map((paragraph, index) =>
        Object.assign({}, paragraph, {
          paraId: (index + 1).toString(16).padStart(8, "0"),
        }),
      );
      const outside: Paragraph = { type: "paragraph", paraId: "00000100", content: [] };
      const document: Document = {
        package: { document: { content: [...wrap(paragraphs, container), outside] } },
      };
      const original = structuredClone(document);
      let trackedDocument = document;
      let directDocument = document;
      const appliedBatches: AppliedDocumentOp[] = [];
      const ops: DocumentOp[] = [];
      for (const [batch, blocks] of batches.entries()) {
        const op = {
          type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
          story: OP_STORIES.MAIN,
          at: { type: after ? "after" : "before", blockId: "00000001" },
          blocks: blocks.map((paragraph, index) =>
            Object.assign({}, paragraph, {
              paraId: (16 + batch * 4 + index).toString(16).padStart(8, "0"),
            }),
          ),
          revision: {
            id: 1000 + batch * 100,
            author: "Reviewer",
            date: `2026-02-03T04:05:0${batch}Z`,
          },
          newIds: {
            revision: Array.from({ length: 16 }, (_, index) => 1001 + batch * 100 + index),
          },
        } satisfies InsertBlocksOp;
        const tracked = apply(trackedDocument, op);
        const direct = apply(trackedDocument, { ...op, revision: undefined });
        const accepted = resolve(tracked, REVISION_DECISIONS.ACCEPT);
        const rejected = resolve(tracked, REVISION_DECISIONS.REJECT);
        expect(accepted.document).toStrictEqual(direct.document); // L1
        expect(rejected.document).toStrictEqual(trackedDocument); // L2
        expect(undo(tracked)).toStrictEqual(trackedDocument); // L4
        expect(undo(accepted)).toStrictEqual(tracked.document);
        expect(undo(rejected)).toStrictEqual(tracked.document);
        expect(apply(structuredClone(trackedDocument), structuredClone(op))).toStrictEqual(tracked); // L5
        const byId = new Map(
          storyParagraphs(tracked.document.package.document).map(({ paragraph }) => [
            paragraph.paraId,
            paragraph,
          ]),
        );
        for (const { paragraph } of storyParagraphs(trackedDocument.package.document)) {
          expect(byId.get(paragraph.paraId)).toBe(paragraph); // L6
        }
        const body = tracked.document.package.document;
        expect(
          storyParagraphs(body).filter(
            (at) => endsItsContainer(body, at) && at.paragraph.pPrMark !== undefined,
          ),
        ).toEqual([]); // L7
        expect(contractViolation(tracked.document)).toBeUndefined();
        expect(
          resolve({ ...tracked, document: accepted.document }, REVISION_DECISIONS.ACCEPT).document,
        ).toBe(accepted.document);
        expect(
          resolve({ ...tracked, document: rejected.document }, REVISION_DECISIONS.REJECT).document,
        ).toBe(rejected.document);
        const demand = revisionIdDemand(trackedDocument, op);
        if (demand.isErr()) throw demand.error;
        expect(demand.value).toBe(tracked.revisions.length - 1);
        trackedDocument = tracked.document;
        directDocument = apply(directDocument, { ...op, revision: undefined }).document;
        appliedBatches.push(tracked);
        ops.push(op);
      }
      // L3: batch acceptance agrees with direct application; rejection agrees
      // with reverse per-operation review and restores every original identity.
      const revisions = appliedBatches.flatMap(({ revisions: recorded }) => recorded);
      const all = {
        document: trackedDocument,
        revisions,
        inverse: [],
        touched: { modified: [], inserted: [], removed: [] },
      };
      expect(resolve(all, REVISION_DECISIONS.ACCEPT).document).toStrictEqual(directDocument);
      expect(resolve(all, REVISION_DECISIONS.REJECT).document).toStrictEqual(document);
      let reversed = trackedDocument;
      for (const applied of appliedBatches.toReversed()) {
        reversed = resolve({ ...applied, document: reversed }, REVISION_DECISIONS.REJECT).document;
      }
      expect(reversed).toStrictEqual(document);
      const batch = applyDocumentOps(document, ops);
      if (batch.isErr()) throw batch.error;
      expect(batch.value.document).toStrictEqual(trackedDocument);
      expect(undo(batch.value)).toStrictEqual(document);
      expect(document).toStrictEqual(original);
    }),
    { numRuns: 10_000 },
  );
});
