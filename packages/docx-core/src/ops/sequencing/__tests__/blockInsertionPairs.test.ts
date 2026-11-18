import { panic } from "better-result";
import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../../test/property-testing";
import type { Document, Paragraph } from "../../../model/document";
import { applyDocumentOps } from "../../apply";
import { equalForStaleness } from "../../equality";
import { formatParaId, type IdLease, type IdLeaseKind } from "../../idLease";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../../refusal";
import {
  DOCUMENT_OP_SCHEMA_VERSION,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  REVISION_DECISIONS,
} from "../../types";
import { applyBatch } from "../applyBatch";
import { createClient } from "../client";
import {
  BATCH_REJECTION_REASONS,
  type DocumentBatch,
  type SequencedBatch,
} from "../envelope";
import { createSequencer } from "../sequencer";
import { transformBatch } from "../transform";

const ANCHOR = "10000005";

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

/** A body whose last paragraph is the anchor, with another before it. */
const base: Document = {
  package: {
    document: {
      content: [paragraph("10000004", "Before"), paragraph(ANCHOR, "Anchor")],
    },
  },
};

const batchOf = (opId: string, op: DocumentOp): DocumentBatch => ({
  schema: DOCUMENT_OP_SCHEMA_VERSION,
  opId,
  actor: opId,
  baseRev: 0,
  ops: [op],
});

const directOp = {
  type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
  story: OP_STORIES.MAIN,
  at: { type: "after", blockId: ANCHOR },
  blocks: [paragraph("00001001", "Direct")],
} as const satisfies DocumentOp;
const direct = batchOf("direct", directOp);

const trackedOp = {
  type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
  story: OP_STORIES.MAIN,
  at: { type: "after", blockId: ANCHOR },
  blocks: [paragraph("00000001", "Tracked")],
  revision: { id: 11, author: "actor-0", date: "2026-01-01T00:00:10.000Z" },
  newIds: {
    revision: [12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25],
    control: [],
  },
} as const satisfies DocumentOp;
const tracked = batchOf("tracked", trackedOp);

const leasedStart = (lease: IdLease, kind: IdLeaseKind): number =>
  lease.blocks.find((block) => block.kind === kind)?.start ??
  panic(`The insertion needs its ${kind} lease.`);

const transformOver = (first: DocumentBatch, second: DocumentBatch) => {
  const applied = applyBatch(base, first.ops);
  if (applied.isErr()) throw applied.error;
  const sequenced: SequencedBatch = {
    ...first,
    revision: 1,
    effects: applied.value.effects,
  };
  return transformBatch(second, [sequenced]);
};

type InOrderOptions = {
  first: DocumentBatch;
  second: DocumentBatch;
  order: "before" | "after";
};

const inOrder = ({ first, second, order }: InOrderOptions): Document => {
  const applied = applyBatch(base, first.ops);
  if (applied.isErr()) throw applied.error;
  const rebased = transformBatch(
    second,
    [{ ...first, revision: 1, effects: applied.value.effects }],
    { order },
  );
  if (rebased.isErr()) throw rebased.error;
  const after = applyBatch(applied.value.document, rebased.value.ops);
  if (after.isErr()) throw after.error;
  return after.value.document;
};

const formattingArbitrary = fc.record(
  {
    keepNext: fc.boolean(),
    indentLeft: fc.integer({ min: 0, max: 1440 }),
  },
  { requiredKeys: [] },
);

const formattingHistory = fc.record({
  anchorFormatting: formattingArbitrary,
  currentFormatting: formattingArbitrary,
  directFormatting: formattingArbitrary,
  insertedFormatting: formattingArbitrary,
  count: fc.integer({ min: 1, max: 3 }),
});

describe("paragraph insertions sharing an anchor", () => {
  test.each([1, 2])(
    "a direct and a tracked insertion of %s paragraphs after one anchor refuse in both directions",
    (count) => {
      const several = batchOf("tracked", {
        ...trackedOp,
        blocks: Array.from({ length: count }, (_, index) =>
          paragraph(formatParaId(index + 1), "Tracked"),
        ),
      });
      for (const [first, second] of [
        [direct, several],
        [several, direct],
      ] as const) {
        const transformed = transformOver(first, second);
        expect(transformed.isErr()).toBe(true);
        if (transformed.isErr())
          expect(transformed.error.reason).toBe(
            BATCH_REJECTION_REASONS.UNSUPPORTED_PAIR,
          );
      }
    },
  );

  test("same-kind, before-anchor and independent insertions still commute", () => {
    const otherDirect = batchOf("other", {
      ...directOp,
      blocks: [paragraph("00001002", "Other")],
    });
    const directBefore = batchOf("direct", {
      ...directOp,
      at: { type: "before", blockId: ANCHOR },
    });
    const trackedBefore = batchOf("tracked", {
      ...trackedOp,
      at: { type: "before", blockId: ANCHOR },
    });
    const independent = batchOf("tracked", {
      ...trackedOp,
      at: { type: "before", blockId: "10000004" },
    });
    for (const [first, second] of [
      [direct, otherDirect],
      [directBefore, trackedBefore],
      [direct, independent],
    ] as const) {
      expect(
        equalForStaleness(
          inOrder({ first, second, order: "after" }),
          inOrder({ first: second, second: first, order: "before" }),
        ),
      ).toBe(true);
    }
  });

  test("a direct insertion cannot transfer a pending mark to differently formatted content", () => {
    const marked = applyDocumentOps(base, tracked.ops);
    if (marked.isErr()) throw marked.error;
    const before = structuredClone(marked.value.document);
    const insertion = {
      ...directOp,
      blocks: [
        { ...paragraph("00001001", "Direct"), formatting: { keepNext: true } },
      ],
    } as const satisfies DocumentOp;
    const applied = applyDocumentOps(marked.value.document, [insertion]);
    expect(applied.isErr()).toBe(true);
    if (applied.isErr())
      expect(applied.error.reason).toBe(
        DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      );
    expect(marked.value.document).toStrictEqual(before);
    for (const [first, second] of [
      [batchOf("direct", insertion), tracked],
      [tracked, batchOf("direct", insertion)],
    ] as const) {
      expect(transformOver(first, second).isErr()).toBe(true);
    }
  });

  test.each(["direct", "tracked"] as const)(
    "clients drop the unsupported peer insertion when %s is sequenced first",
    (firstKind) => {
      const authority = createSequencer(base);
      const directLease = authority.grantIdLease("direct");
      if (directLease.isErr()) throw directLease.error;
      const trackedLease = authority.grantIdLease("tracked");
      if (trackedLease.isErr()) throw trackedLease.error;
      const directBatch = batchOf("direct", {
        ...directOp,
        blocks: [
          {
            ...paragraph(
              formatParaId(leasedStart(directLease.value, "paragraph")),
              "Direct",
            ),
            formatting: { keepNext: true },
          },
        ],
      });
      const revision = leasedStart(trackedLease.value, "revision");
      const trackedBatch = batchOf("tracked", {
        ...trackedOp,
        blocks: [
          paragraph(
            formatParaId(leasedStart(trackedLease.value, "paragraph")),
            "Tracked",
          ),
        ],
        revision: { ...trackedOp.revision, id: revision },
        newIds: { revision: [revision + 1, revision + 2] },
      });
      const [first, second] =
        firstKind === "direct"
          ? [directBatch, trackedBatch]
          : [trackedBatch, directBatch];
      if (first === undefined || second === undefined)
        panic("Both insertion batches are stated.");
      const firstPeer = createClient(base);
      const secondPeer = createClient(base);
      expect(firstPeer.enqueue(first).isOk()).toBe(true);
      expect(secondPeer.enqueue(second).isOk()).toBe(true);
      const previous = secondPeer.document;
      const submission = authority.submit(first);
      if (submission.type !== "ack") throw submission.reason;
      firstPeer.receiveAck(submission);
      const broadcast = authority.broadcasts.at(0);
      if (broadcast === undefined) panic("The first insertion was sequenced.");
      expect(firstPeer.receiveBroadcast(broadcast).isOk()).toBe(true);
      const changes = secondPeer.receiveBroadcast(broadcast);
      if (changes.isErr()) throw changes.error;
      expect(changes.value.map(({ type }) => type)).toEqual(["replaced"]);
      const replacement = changes.value.at(0);
      if (replacement === undefined)
        panic("The pending insertion was dropped.");
      const replaced = applyBatch(previous, replacement.batch.ops);
      if (replaced.isErr()) throw replaced.error;
      expect(
        equalForStaleness(replaced.value.document, authority.document),
      ).toBe(true);
      const rejected = authority.submit(second);
      if (rejected.type !== "reject") panic("The mixed insertion must refuse.");
      expect(rejected.reason.reason).toBe(
        BATCH_REJECTION_REASONS.UNSUPPORTED_PAIR,
      );
      secondPeer.receiveReject(rejected);
      expect(secondPeer.pending).toEqual([]);
      expect(secondPeer.notices.map(({ reason }) => reason.reason)).toEqual([
        BATCH_REJECTION_REASONS.UNSUPPORTED_PAIR,
      ]);
      for (const peer of [firstPeer, secondPeer])
        expect(equalForStaleness(peer.document, authority.document)).toBe(true);
    },
  );

  test(
    "sequential insertions after a pending mark refuse without changing any formatting history",
    () => {
      assertProperty(
        fc.property(formattingHistory, (history) => {
          const formattedBase = {
            package: {
              document: {
                content: [
                  paragraph("10000004", "Before"),
                  {
                    ...paragraph(ANCHOR, "Anchor"),
                    formatting: history.anchorFormatting,
                  },
                ],
              },
            },
          } satisfies Document;
          const marked = applyDocumentOps(formattedBase, [
            {
              ...trackedOp,
              blocks: Array.from({ length: history.count }, (_, index) => ({
                ...paragraph(formatParaId(index + 1), "Tracked"),
                formatting: history.insertedFormatting,
              })),
            },
          ]);
          if (marked.isErr()) throw marked.error;
          const changed = applyDocumentOps(marked.value.document, [
            {
              type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
              story: OP_STORIES.MAIN,
              blockId: ANCHOR,
              patch: { keepNext: false, ...history.currentFormatting },
            },
          ]);
          if (changed.isErr()) throw changed.error;
          const before = structuredClone(changed.value.document);
          const inserted = applyDocumentOps(changed.value.document, [
            {
              ...directOp,
              blocks: [
                {
                  ...paragraph("00001001", "Direct"),
                  formatting: history.directFormatting,
                },
              ],
            },
          ]);
          expect(inserted.isErr()).toBe(true);
          if (inserted.isErr())
            expect(inserted.error.reason).toBe(
              DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
            );
          expect(changed.value.document).toStrictEqual(before);
        }),
        {
          id: "pending paragraph mark insertion preserves formatting history",
          numRuns: 80,
          examples: [
            [
              {
                anchorFormatting: {},
                currentFormatting: { keepNext: true },
                directFormatting: { keepNext: true },
                insertedFormatting: {},
                count: 1,
              },
            ],
          ],
        },
      );
    },
    propertyTestTimeout(30_000),
  );

  test.each([REVISION_DECISIONS.ACCEPT, REVISION_DECISIONS.REJECT])(
    "an accepted direct insertion before the pending mark keeps its paragraph and formatting after %s",
    (decision) => {
      const marked = applyDocumentOps(base, tracked.ops);
      if (marked.isErr()) throw marked.error;
      const directParagraph = {
        ...paragraph("00001001", "Direct"),
        formatting: { keepNext: true },
      } satisfies Paragraph;
      const inserted = applyDocumentOps(marked.value.document, [
        {
          ...directOp,
          at: { type: "before", blockId: ANCHOR },
          blocks: [directParagraph],
        },
      ]);
      if (inserted.isErr()) throw inserted.error;
      const resolved = applyDocumentOps(inserted.value.document, [
        {
          type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
          story: OP_STORIES.MAIN,
          revisionIds: marked.value.revisions,
          decision,
        },
      ]);
      if (resolved.isErr()) throw resolved.error;
      const expected = {
        package: {
          document: {
            content: [
              paragraph("10000004", "Before"),
              directParagraph,
              paragraph(ANCHOR, "Anchor"),
              ...(decision === REVISION_DECISIONS.ACCEPT
                ? trackedOp.blocks
                : []),
            ],
          },
        },
      } satisfies Document;
      expect(equalForStaleness(resolved.value.document, expected)).toBe(true);
    },
  );
});
