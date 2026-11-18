import { expect, test } from "bun:test";
import fc from "fast-check";
import { EditorState } from "prosemirror-state";

import {
  type IdLease,
  idLeaseHolds,
  introducedIds,
  leasedIdCensus,
  normalizeForOps,
} from "@stll/docx-core/ops";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { createFolioAIEditSnapshot } from "../ai-edits/snapshot";
import { CanonicalPublicOperations } from "../canonical/canonicalPublicOperations";
import {
  type CanonicalCommit,
  createCanonicalSession,
  publishCanonicalProjection,
} from "../canonical/canonicalSession";
import type { FolioDocumentOperation } from "../document-operations";

const paragraph = (paraId: string) => ({
  type: "paragraph" as const,
  paraId,
  content: [
    {
      type: "run" as const,
      content: [{ type: "text" as const, text: paraId }],
    },
  ],
});

type InsertionOptions = {
  id: string;
  blockId: string;
  side: "before" | "after";
  comment: "attached" | "none";
};
const paragraphInsertion = ({
  id,
  blockId,
  side,
  comment,
  text,
}: InsertionOptions & { text: string }) =>
  ({
    id,
    type: side === "before" ? "insertBeforeBlock" : "insertAfterBlock",
    blockId,
    text,
    ...(comment === "attached" ? { comment: { text: `${id} comment` } } : {}),
  }) satisfies FolioDocumentOperation;

const INSERTIONS = {
  paragraph: (options: InsertionOptions) =>
    paragraphInsertion({ ...options, text: options.id }),
  paragraphs: (options: InsertionOptions) =>
    paragraphInsertion({ ...options, text: `${options.id}\nExtra` }),
  table: ({ id, blockId, side }: InsertionOptions) =>
    ({
      id,
      type: "insertTable",
      blockId,
      position: side,
      rows: [[`${id} term`, `${id} value`]],
    }) satisfies FolioDocumentOperation,
};

const leaseBounds = fc.record({
  paragraphs: fc.integer({ min: 0, max: 12 }),
  revisions: fc.integer({ min: 0, max: 40 }),
  controls: fc.integer({ min: 0, max: 12 }),
  comments: fc.integer({ min: 0, max: 2 }),
});
const AMPLE = { paragraphs: 12, revisions: 40, controls: 12, comments: 2 };

test(
  "insertion-gap replay stays within admitted leases and undoes atomically",
  async () => {
    await assertProperty(
      fc.asyncProperty(leaseBounds, async (scarce) => {
        // Exercise every declared producer and both source topologies for each
        // lease; ample controls ensure exhausted leases cannot hide replay.
        for (const mode of ["direct", "tracked-changes"] as const)
          for (const bounds of [scarce, AMPLE])
            for (const terminal of [false, true])
              for (const comment of ["none", "attached"] as const)
                for (const first of Object.values(INSERTIONS))
                  for (const second of Object.values(INSERTIONS)) {
                    const source = normalizeForOps({
                      package: {
                        document: {
                          content: terminal
                            ? [paragraph("10000001")]
                            : [paragraph("10000001"), paragraph("10000002")],
                        },
                      },
                    });
                    const original = structuredClone(source);
                    const lease = {
                      blocks: (
                        [
                          {
                            kind: "paragraph",
                            start: 0x20000001,
                            end: 0x20000001 + bounds.paragraphs,
                          },
                          {
                            kind: "revision",
                            start: 1,
                            end: 1 + bounds.revisions,
                          },
                          {
                            kind: "control",
                            start: 1,
                            end: 1 + bounds.controls,
                          },
                          {
                            kind: "comment",
                            start: 1,
                            end: 1 + bounds.comments,
                          },
                        ] as const
                      ).filter(({ start, end }) => start < end),
                    } satisfies IdLease;
                    const leaseBefore = structuredClone(lease);
                    const session = createCanonicalSession(source, undefined, {
                      idLease: lease,
                    }).unwrap();
                    let state = EditorState.create({
                      doc: session.projection.doc,
                    });
                    const publish = (commit: CanonicalCommit) => {
                      const published = publishCanonicalProjection({
                        state,
                        session,
                        commit,
                      });
                      if (published.isErr()) return false;
                      state = published.value.state;
                      return true;
                    };
                    const executor = new CanonicalPublicOperations({
                      session,
                      getState: () => state,
                      publish,
                    });
                    const operations = [
                      first({
                        id: "first",
                        blockId: "10000001",
                        side: "after",
                        comment,
                      }),
                      second({
                        id: "second",
                        blockId: terminal ? "10000001" : "10000002",
                        side: terminal ? "after" : "before",
                        comment,
                      }),
                    ];
                    const result = executor.apply({
                      snapshot: createFolioAIEditSnapshot(state.doc),
                      batch: { version: 1, mode, operations },
                      undefinedReferences: "refuse",
                      revisionStamp: {
                        idSeed: 1,
                        date: "2026-01-02T03:04:05Z",
                      },
                    });
                    const accounted = [...result.applied, ...result.skipped]
                      .map(({ id }) => id)
                      .toSorted();
                    expect(accounted).toEqual(
                      operations.map(({ id }) => id).toSorted(),
                    );
                    if (
                      bounds === AMPLE &&
                      !terminal &&
                      (comment === "none" || mode === "direct")
                    ) {
                      expect(result.skipped).toEqual([]);
                      expect(result.applied).toHaveLength(2);
                    }
                    const introduced = introducedIds(
                      leasedIdCensus(original),
                      leasedIdCensus(session.document),
                    );
                    for (const { kind, id } of introduced)
                      expect(idLeaseHolds(lease, kind, id)).toBe(true);
                    const revisions = result.applied.flatMap(
                      ({ revisionIds }) => revisionIds ?? [],
                    );
                    expect(new Set(revisions).size).toBe(revisions.length);
                    expect(lease).toEqual(leaseBefore);
                    expect(source).toEqual(original);
                    const edited = structuredClone(session.document);
                    if (result.applied.length === 0) {
                      expect(session.document).toEqual(original);
                      expect(session.version).toBe(0);
                      expect(session.canUndo).toBe(false);
                      continue;
                    }
                    expect(session.version).toBe(1);
                    expect(publish(session.prepareUndo(state).unwrap())).toBe(
                      true,
                    );
                    expect(session.document).toEqual(original);
                    expect(publish(session.prepareRedo(state).unwrap())).toBe(
                      true,
                    );
                    expect(session.document).toEqual(edited);
                  }
      }),
      {
        id: "insertion-gap replay stays within admitted leases and undoes atomically",
        numRuns: 12,
      },
    );
  },
  propertyTestTimeout(30_000),
);
