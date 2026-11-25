import { panic } from "better-result";
import { expect, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { createDocx } from "../docx/rezip";
import { paragraph } from "../docx/server/build";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";
import type { FolioPendingSuggestionRecord } from "./pending-suggestions";
import { hashFolioAIBlockContent } from "./snapshot";

// A resolution can remove the paragraph a pending suggestion's record names:
// a rejected split or insertion joins it into its neighbour, an accepted
// deletion of its mark joins it into the next. The record then moves to the
// saved paragraph that keeps the suggestion's words or place, or is retired,
// and a reload reports it stale with that reason.

const SOURCE = [
  ["10000001", "Opening clause."],
  ["10000002", "Signed in two copies."],
  ["10000003", "Closing clause."],
] as const;

const source = () => {
  const document = createEmptyDocument();
  document.package.document.content = SOURCE.map(([paraId, text]) =>
    Object.assign(paragraph(text), { paraId }),
  );
  return createDocx(document);
};

// Mostly one paragraph, so the batches meet at its marks, and the first
// paragraphs a batch adds (their ids are minted in order).
const blockId = fc.oneof(
  { weight: 3, arbitrary: fc.constant("10000002") },
  { weight: 1, arbitrary: fc.constantFrom("10000001", "10000003") },
  { weight: 2, arbitrary: fc.constantFrom("00000001", "00000002") },
);
const words = fc.constantFrom("Payment due.", "New term", "Amended text.");
const operation: fc.Arbitrary<FolioDocumentOperation> = fc.oneof(
  {
    weight: 3,
    arbitrary: fc.record({
      type: fc.constant("splitBlock" as const),
      blockId,
      offset: fc.integer({ min: 1, max: 3 }),
    }),
  },
  {
    weight: 3,
    arbitrary: fc.record({
      type: fc.constantFrom(
        "insertAfterBlock" as const,
        "insertBeforeBlock" as const,
      ),
      blockId,
      text: words,
    }),
  },
  {
    weight: 1,
    arbitrary: fc.record({
      type: fc.constant("replaceBlock" as const),
      blockId,
      text: words,
    }),
  },
  {
    weight: 1,
    arbitrary: fc.record({
      type: fc.constant("deleteBlock" as const),
      blockId,
    }),
  },
  {
    weight: 1,
    arbitrary: fc.record({
      type: fc.constant("commentOnBlock" as const),
      blockId,
      comment: fc.record({ text: words }),
    }),
  },
) as fc.Arbitrary<FolioDocumentOperation>;

const batch = fc.record({
  mode: fc.constantFrom("tracked-changes" as const, "suggested" as const),
  operation,
});

const resolution = fc.record({
  verdict: fc.constantFrom("accept" as const, "reject" as const),
  target: fc.constantFrom(
    "change" as const,
    "suggestion" as const,
    "all" as const,
  ),
  pick: fc.nat(),
});

type Resolution = typeof resolution extends fc.Arbitrary<infer T> ? T : never;

type ResolveOptions = {
  reviewer: FolioDocxReviewer;
  /** Every suggestion the batches staged, resolved or not. */
  suggestionIds: readonly string[];
  step: Resolution;
};

const resolve = ({
  reviewer,
  suggestionIds,
  step: { verdict, target, pick },
}: ResolveOptions) => {
  switch (target) {
    case "change": {
      const changes = reviewer.getChanges();
      const change = changes[pick % Math.max(changes.length, 1)];
      if (change === undefined) return;
      if (verdict === "accept") reviewer.acceptChange(change);
      else reviewer.rejectChange(change);
      return;
    }
    case "suggestion": {
      const id = suggestionIds[pick % Math.max(suggestionIds.length, 1)];
      if (id === undefined) return;
      if (verdict === "accept") reviewer.acceptSuggestion(id);
      else reviewer.rejectSuggestion(id);
      return;
    }
    case "all":
      if (verdict === "accept") reviewer.acceptAll();
      else reviewer.rejectAll();
      return;
    default:
      return target satisfies never;
  }
};

/** Whether each record names a paragraph the story shows, or is retired. */
const anchoring = (
  reviewer: FolioDocxReviewer,
  records: readonly FolioPendingSuggestionRecord[],
) => {
  const { anchors } = reviewer.snapshot();
  return records.map(({ suggestionId, anchor: { paraId }, retired }) => ({
    suggestionId,
    anchored:
      retired !== undefined ||
      (paraId !== null && anchors[paraId] !== undefined),
  }));
};

test(
  "a pending suggestion's record names a shown paragraph, or is retired, after any resolution",
  async () => {
    const bytes = await source();
    await assertProperty(
      fc.asyncProperty(
        fc.array(batch, { minLength: 2, maxLength: 6 }),
        fc.array(resolution, { minLength: 1, maxLength: 8 }),
        async (batches, resolutions) => {
          const reviewer = await FolioDocxReviewer.fromBuffer(bytes.slice(0));
          const suggestionIds: string[] = [];
          for (const [index, { mode, operation: op }] of batches.entries()) {
            const { applied } = reviewer.applyDocumentOperations(
              {
                version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
                mode,
                operations: [{ ...op, id: `op-${String(index)}` }],
              },
              { undefinedReferences: "refuse" },
            );
            for (const { suggestionId } of applied)
              if (suggestionId !== undefined) suggestionIds.push(suggestionId);
          }
          // Each resolution re-anchors the records whose paragraph it removed.
          for (const step of resolutions)
            resolve({ reviewer, suggestionIds, step });
          const records = reviewer.exportPendingSuggestions();
          const anchored = anchoring(reviewer, records);
          expect(anchored).toEqual(
            anchored.map(({ suggestionId }) => ({
              suggestionId,
              anchored: true,
            })),
          );
          // A retired record reports why it no longer restages.
          const reopened = await FolioDocxReviewer.fromBuffer(
            await reviewer.toBuffer(),
          );
          const loaded = reopened.loadPendingSuggestions(records);
          const retired = records.flatMap(
            ({ suggestionId, retired: reason }) =>
              reason === undefined ? [] : [{ suggestionId, reason }],
          );
          expect(
            loaded.filter((_, index) => records[index]?.retired !== undefined),
          ).toEqual(
            retired.map(({ suggestionId, reason }) => ({
              status: "stale",
              suggestionId,
              reason,
            })),
          );
        },
      ),
      { id: "suggestion-anchor-after-resolution", numRuns: 300 },
    );
  },
  propertyTestTimeout(120_000),
);

test("export rebases the insertion a cancelled tracked paragraph anchored", async () => {
  const reviewer = await FolioDocxReviewer.fromBuffer(await source());
  const apply = (
    mode: "suggested" | "tracked-changes",
    request: FolioDocumentOperation,
  ) => {
    const result = reviewer.applyDocumentOperations(
      {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode,
        operations: [request],
      },
      { undefinedReferences: "refuse" },
    );
    expect(result.applied).toHaveLength(1);
    expect(result.skipped).toEqual([]);
  };

  apply("suggested", {
    id: "first",
    type: "insertAfterBlock",
    blockId: "10000002",
    text: "Earlier proposal.",
  });
  const existing = new Set(reviewer.getContent().map(({ id }) => id));
  const trackedText = "Tracked paragraph.";
  apply("tracked-changes", {
    id: "tracked",
    type: "insertAfterBlock",
    blockId: "10000002",
    text: trackedText,
  });
  const tracked = reviewer.getContent().filter(({ id }) => !existing.has(id));
  expect(tracked).toHaveLength(1);
  const inserted = tracked.at(0) ?? panic("The tracked paragraph is missing.");
  apply("suggested", {
    id: "second",
    type: "insertAfterBlock",
    blockId: inserted.id,
    text: "Later proposal.",
    precondition: { blockTextHash: hashFolioAIBlockContent(trackedText) },
  });
  apply("tracked-changes", {
    id: "cancel",
    type: "deleteBlock",
    blockId: inserted.id,
  });
  // The deletion has already taken the tracked insertion out of the document.
  expect(reviewer.snapshot().anchors[inserted.id]).toBeUndefined();
  expect(reviewer.acceptAll().resolved).toBe(0);

  const records = reviewer.exportPendingSuggestions();
  expect(records.map(({ suggestionId }) => suggestionId)).toEqual([
    "first",
    "second",
  ]);
  expect(reviewer.exportPendingSuggestions()).toEqual(records);
  const reopened = await FolioDocxReviewer.fromBuffer(
    await reviewer.toBuffer(),
  );
  expect(reopened.loadPendingSuggestions(records)).toEqual([
    { status: "restaged", suggestionId: "first" },
    { status: "restaged", suggestionId: "second" },
  ]);
  expect(reopened.getContent().map(({ text }) => text)).toEqual(
    reviewer.getContent().map(({ text }) => text),
  );
});
