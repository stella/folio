import { panic } from "better-result";
import { expect, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../../test/property-testing";
import { serializeDocumentToDocx } from "../../../serialize/docx";
import { createEditorIntentIdAllocator } from "../../editorIntent";
import { equalForStaleness } from "../../equality";
import { DOCUMENT_OP_SCHEMA_VERSION, DOCUMENT_OP_TYPES } from "../../types";
import { applyBatch } from "../applyBatch";
import { BATCH_REJECTION_REASONS, type DocumentBatch } from "../envelope";
import { createSequencer } from "../sequencer";
import { transformBatch } from "../transform";
import { canonicalBase, compileCanonicalEdit } from "./canonicalEdits";

type CommentDecisionOptions = {
  kind: "comment" | "pointComment";
  offset: number;
  span: number;
  placement: "independent" | "insideRevision";
  decision: "accept" | "reject";
};

const commentDecision = ({
  kind,
  offset,
  span,
  placement,
  decision,
}: CommentDecisionOptions) => {
  const document = canonicalBase();
  const sequencer = createSequencer(document);
  const author = {
    actor: "commenter",
    tag: "commenter",
    lease: sequencer.grantIdLease("commenter").unwrap(),
    allocate: createEditorIntentIdAllocator(),
  };
  const comment = compileCanonicalEdit(
    document,
    author,
    {
      kind,
      // Seed 44 addresses offset 6 inside canonicalBase's revision 10.
      // The control below verifies its actual anchor and revision allocation.
      a: placement === "insideRevision" ? 44 : offset,
      b: span,
      ...(placement === "independent" ? { focus: true } : {}),
      suggest: false,
    },
    20,
  );
  const resolve = compileCanonicalEdit(
    document,
    author,
    { kind: "decide", a: 0, b: decision === "accept" ? 0 : 1, suggest: false },
    21,
  );
  if (comment === undefined || resolve === undefined)
    panic("The actual comment and decision producers must compile.");
  const commentBatch = {
    schema: DOCUMENT_OP_SCHEMA_VERSION,
    opId: "comment",
    actor: author.actor,
    baseRev: 0,
    ops: comment,
  } satisfies DocumentBatch;
  const decisionBatch = {
    schema: DOCUMENT_OP_SCHEMA_VERSION,
    opId: "decision",
    actor: "decider",
    baseRev: 0,
    ops: resolve,
  } satisfies DocumentBatch;
  const creation = comment.find(
    (op) => op.type === DOCUMENT_OP_TYPES.CREATE_COMMENT,
  );
  if (creation === undefined)
    panic("The comment producer must create a comment.");
  return { document, commentBatch, decisionBatch, creation };
};

test(
  "revision decisions commute with independently anchored comments in either order",
  async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.integer({ min: 0, max: 15 }),
        fc.integer({ min: 0, max: 2 }),
        async (offset, span) => {
          for (const kind of ["comment", "pointComment"] as const)
            for (const decision of ["accept", "reject"] as const) {
              const { document, commentBatch, decisionBatch, creation } =
                commentDecision({
                  kind,
                  offset,
                  span,
                  decision,
                  placement: "independent",
                });
              expect(creation.newIds?.revision).toEqual([]);
              const commented = applyBatch(document, commentBatch.ops).unwrap();
              const resolved = applyBatch(document, decisionBatch.ops).unwrap();
              const overComment = transformBatch(decisionBatch, [
                { ...commentBatch, revision: 1, effects: commented.effects },
              ]).unwrap();
              const overDecision = transformBatch(commentBatch, [
                { ...decisionBatch, revision: 1, effects: resolved.effects },
              ]).unwrap();
              expect(overComment.ops).toEqual(decisionBatch.ops);
              expect(overDecision.ops).toEqual(commentBatch.ops);
              const left = applyBatch(
                commented.document,
                overComment.ops,
              ).unwrap().document;
              const right = applyBatch(
                resolved.document,
                overDecision.ops,
              ).unwrap().document;
              const expected = applyBatch(document, [
                ...decisionBatch.ops,
                ...commentBatch.ops,
              ]).unwrap().document;
              expect(equalForStaleness(left, expected)).toBe(true);
              expect(equalForStaleness(right, expected)).toBe(true);
              expect(
                new Uint8Array(await serializeDocumentToDocx(left)),
              ).toEqual(new Uint8Array(await serializeDocumentToDocx(right)));
            }
        },
      ),
      {
        id: "revision decisions commute with independently anchored comments",
        numRuns: 18,
      },
    );
  },
  propertyTestTimeout(10_000),
);

for (const kind of ["comment", "pointComment"] as const)
  for (const decision of ["accept", "reject"] as const)
    test(`${decision} refuses a ${kind} producer that cuts the decided revision`, () => {
      const { document, commentBatch, decisionBatch, creation } =
        commentDecision({
          kind,
          offset: 0,
          span: 0,
          decision,
          placement: "insideRevision",
        });
      expect(creation.anchor).toMatchObject(
        kind === "pointComment"
          ? { kind: "point", at: { blockId: "10000011", offset: 6 } }
          : { kind: "range", from: { blockId: "10000011", offset: 6 } },
      );
      expect(creation.newIds?.revision?.length).toBeGreaterThan(0);
      const commented = applyBatch(document, commentBatch.ops).unwrap();
      const resolved = applyBatch(document, decisionBatch.ops).unwrap();
      const left = applyBatch(commented.document, decisionBatch.ops).unwrap();
      const right = applyBatch(resolved.document, commentBatch.ops).unwrap();
      // The producer gave a cut piece a new revision id: a decision of only
      // the original id cannot commute across that cut.
      expect(equalForStaleness(left.document, right.document)).toBe(false);
      const transformed = transformBatch(decisionBatch, [
        { ...commentBatch, revision: 1, effects: commented.effects },
      ]);
      expect(transformed.isErr()).toBe(true);
      if (transformed.isErr())
        expect(transformed.error).toMatchObject({
          reason: BATCH_REJECTION_REASONS.UNSUPPORTED_PAIR,
          opType: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
          overType: DOCUMENT_OP_TYPES.CREATE_COMMENT,
        });
    });
