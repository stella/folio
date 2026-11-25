/**
 * A merge's separator goes in where words follow the joined paragraph once
 * every change is accepted, whatever else the batch deletes: applied
 * directly, as tracked changes accepted, and as suggestions accepted, which
 * read the same once saved and reopened.
 *
 * The story holds four paragraphs. A batch merges one of the first three with
 * the next, with a separator, and deletes up to three paragraphs after it, in
 * any order around the merge; each deletion may be refused, as one whose
 * precondition no longer holds is. An earlier tracked batch may already have
 * deleted the last paragraph.
 *
 * A case applies a batch to a fresh package in each mode, then saves and
 * reopens it, about 200 ms; the budget is 60 s for the stated runs, scaled by
 * the nightly factor.
 */

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
import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { FolioDocxReviewer } from "./headless";
import { hashFolioAIBlockText } from "./snapshot";

type Mode = "direct" | "tracked-changes" | "suggested";
type Case = {
  /** The paragraph the merge joins with the next, by index. */
  merged: number;
  separator: string;
  /** The paragraphs after it the batch deletes, by index, and whether refused. */
  deleted: { index: number; refused: boolean }[];
  /** Where the merge stands among the batch's operations. */
  mergeAt: number;
  /** An earlier tracked batch deleted the last paragraph. */
  lastPending: boolean;
};

const TEXTS = [
  "First clause.",
  "Second clause.",
  "Third clause.",
  "Last clause.",
] as const;

/** A text hash no paragraph has: the deletion naming it is refused. */
const STALE_HASH = hashFolioAIBlockText("No paragraph reads this.");

const open = async (): Promise<FolioDocxReviewer> => {
  const bytes = (
    await ensureParaIds(await createDocx(fromMarkdown(TEXTS.join("\n\n"))))
  ).docx;
  return FolioDocxReviewer.fromBuffer(bytes, { author: "Reviewer" });
};

/** Apply `operations` in `mode`, accepting what it suggests as tracked changes. */
const apply = (
  reviewer: FolioDocxReviewer,
  mode: Mode,
  operations: FolioDocumentOperation[],
): string[] => {
  const result = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode,
      atomic: false,
      operations,
    },
    { undefinedReferences: "refuse" },
  );
  for (const suggestionId of new Set(
    result.applied.flatMap(({ suggestionId: id }) => id ?? []),
  ))
    expect(reviewer.acceptSuggestion(suggestionId)).toBe(true);
  return result.applied.map(({ id }) => id).toSorted();
};

/** The paragraphs a reader lists once the batch is applied, accepted and saved. */
const replay = async (setup: Case, mode: Mode) => {
  const reviewer = await open();
  const ids = reviewer.getContent().map(({ id }) => id);
  const idOf = (index: number): string => {
    const id = ids[index];
    if (id === undefined) throw new Error(`Missing paragraph ${index}`);
    return id;
  };
  if (setup.lastPending)
    apply(reviewer, mode === "direct" ? "direct" : "tracked-changes", [
      { id: "delete-last", type: "deleteBlock", blockId: idOf(3) },
    ]);
  const operations: FolioDocumentOperation[] = setup.deleted.map(
    ({ index, refused }) => ({
      id: `delete-${String(index)}`,
      type: "deleteBlock",
      blockId: idOf(index),
      ...(refused ? { precondition: { blockTextHash: STALE_HASH } } : {}),
    }),
  );
  operations.splice(Math.min(setup.mergeAt, operations.length), 0, {
    id: "merge",
    type: "mergeBlockWithNext",
    blockId: idOf(setup.merged),
    separator: setup.separator,
  });
  const applied = apply(reviewer, mode, operations);
  reviewer.acceptAll();
  const saved = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer(), {
    author: "Reviewer",
  });
  return {
    applied,
    texts: saved.getContent().map(({ text }) => text),
  };
};

const caseArbitrary: fc.Arbitrary<Case> = fc
  .record({
    merged: fc.nat({ max: 2 }),
    separator: fc.constantFrom(" ", "; "),
    deletions: fc.subarray([1, 2, 3]),
    refusals: fc.array(fc.boolean(), { minLength: 3, maxLength: 3 }),
    mergeAt: fc.nat({ max: 3 }),
    lastPending: fc.boolean(),
  })
  .map(({ merged, separator, deletions, refusals, mergeAt, lastPending }) => ({
    merged,
    separator,
    // The paragraph pending deletion is not deleted again.
    deleted: deletions
      .filter((index) => index > merged && !(lastPending && index === 3))
      .map((index) => ({ index, refused: refusals[index - 1] === true })),
    mergeAt,
    lastPending,
  }));

const examples: [Case][] = [
  // #1757 seed row b3edf73868e5c893: the batch asks to delete a paragraph
  // after the merge and refuses it, so its words still follow the separator.
  [
    {
      merged: 0,
      separator: " ",
      deleted: [
        { index: 1, refused: false },
        { index: 2, refused: true },
      ],
      mergeAt: 0,
      lastPending: true,
    },
  ],
];

test(
  "a merge separates the words that follow it once accepted in every mode",
  async () => {
    await assertProperty(
      fc.asyncProperty(caseArbitrary, async (setup) => {
        const direct = await replay(setup, "direct");
        for (const mode of ["tracked-changes", "suggested"] as const) {
          const reviewed = await replay(setup, mode);
          // Where one mode refuses what the other applies, the two documents
          // part, and comparing them says nothing about the separator.
          if (
            JSON.stringify(reviewed.applied) !== JSON.stringify(direct.applied)
          )
            continue;
          expect(reviewed.texts).toEqual(direct.texts);
        }
      }),
      {
        id: "a merge separates the words that follow it once accepted in every mode",
        numRuns: 40,
        examples,
      },
    );
  },
  propertyTestTimeout(60_000),
);
