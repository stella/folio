import assert from "node:assert/strict";
import { test } from "node:test";

import { fromMarkdown } from "@stll/folio-core/markdown";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "@stll/folio-core/server";

import { sequentialGroups } from "../support/batch-replay-order.ts";
import { openReviewer, packDocument } from "../support/documents.ts";
import { type Row, rowsOf } from "../support/oracle.ts";

type Reviewer = Awaited<ReturnType<typeof openReviewer>>;
type Mode = "direct" | "tracked-changes";

const SPLIT_OFFSET = "Closing ".length;

/** The paragraph before the split one, of each kind a merge can carry over. */
const BEFORE_SPLIT = [
  "Plain clause.",
  "1. Numbered clause.",
  "## Heading clause",
] as const;
const FINAL = "Closing remarks.";

type PredecessorEdit = (
  rows: readonly Row[],
  predecessor: Row,
) => FolioDocumentOperation;

/**
 * The edits that put a pending property change on the first half of the
 * split, the paragraph before the final one: a merge of the paragraph before
 * it into it, a property write, and a restyle.
 */
const PREDECESSOR_EDITS = {
  merge: (rows, predecessor) => {
    const before = rows.at(rows.indexOf(predecessor) - 1);
    assert.ok(before);
    return {
      id: "edit",
      type: "mergeBlockWithNext",
      blockId: before.id,
      separator: " ",
    };
  },
  properties: (_rows, predecessor) => ({
    id: "edit",
    type: "setBlockParagraphProperties",
    blockId: predecessor.id,
    properties: { alignment: "center" },
  }),
  restyle: (_rows, predecessor) => ({
    id: "edit",
    type: "replaceBlock",
    blockId: predecessor.id,
    text: predecessor.text,
    styleId: "Heading2",
  }),
} as const satisfies Record<string, PredecessorEdit>;

const apply = (
  reviewer: Reviewer,
  mode: Mode,
  operations: FolioDocumentOperation[],
) =>
  reviewer.applyDocumentOperations(
    { version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION, mode, operations },
    { undefinedReferences: "refuse" },
  );

/** `reviewer` reopened from its save with every change resolved one way. */
const resolved = async (
  reviewer: Reviewer,
  resolution: "accept" | "reject",
): Promise<Reviewer> => {
  const copy = await openReviewer(new Uint8Array(await reviewer.toBuffer()));
  if (resolution === "accept") copy.acceptAll();
  else copy.rejectAll();
  return copy;
};

/** What resolving every change one way leaves of the body, ids aside. */
const resolvedView = async (
  reviewer: Reviewer,
  resolution: "accept" | "reject",
) =>
  rowsOf(await resolved(reviewer, resolution)).map(
    ({ id: _id, ...row }) => row,
  );

/** `fixture` with its final paragraph split, in `mode`. */
const splitFinal = async (fixture: Uint8Array, mode: Mode) => {
  const reviewer = await openReviewer(fixture);
  const final = rowsOf(reviewer).at(-1);
  assert.ok(final);
  assert.equal(
    apply(reviewer, mode, [
      {
        id: "split",
        type: "splitBlock",
        blockId: final.id,
        offset: SPLIT_OFFSET,
      },
    ]).applied.length,
    1,
  );
  return reviewer;
};

type BatchOptions = {
  reviewer: Reviewer;
  edit: PredecessorEdit;
  deletionFirst: boolean;
};

/** `edit` and the deletion of the split's second half, against `reviewer`'s live rows. */
const batchOf = ({
  reviewer,
  edit,
  deletionFirst,
}: BatchOptions): FolioDocumentOperation[] => {
  const rows = rowsOf(reviewer);
  const [predecessor, final] = rows.slice(-2);
  assert.ok(predecessor && final);
  const deletion: FolioDocumentOperation = {
    id: "delete",
    type: "deleteBlock",
    blockId: final.id,
  };
  const editOperation = edit(rows, predecessor);
  return deletionFirst ? [deletion, editOperation] : [editOperation, deletion];
};

// #1356: a tracked batch merging into the first half of a split and deleting
// the second half kept the emptied second half. The merge's pending property
// change on the first half refused retracting the split's inserted break,
// which one at a time the deletion retracts before the merge runs.
test("deleting a split's second half retracts the break whatever pending property change the batch puts before it", async () => {
  for (const kind of BEFORE_SPLIT) {
    const fixture = await packDocument(
      fromMarkdown(["Opening clause.", kind, FINAL].join("\n\n")),
    );
    const original = await resolvedView(await openReviewer(fixture), "accept");
    for (const [name, edit] of Object.entries(PREDECESSOR_EDITS)) {
      for (const deletionFirst of [false, true]) {
        const context = `${kind} ${name}, ${deletionFirst ? "deletion first" : "edit first"}`;
        const direct = await splitFinal(fixture, "direct");
        const directOperations = batchOf({
          reviewer: direct,
          edit,
          deletionFirst,
        });
        assert.equal(
          apply(direct, "direct", directOperations).applied.length,
          2,
          context,
        );
        const expected = await resolvedView(direct, "accept");
        assert.equal(
          expected.length,
          original.length - (name === "merge" ? 1 : 0),
          `${context}: directly, the deletion takes the second half away`,
        );

        const batch = await splitFinal(fixture, "tracked-changes");
        assert.equal(
          apply(
            batch,
            "tracked-changes",
            batchOf({ reviewer: batch, edit, deletionFirst }),
          ).applied.length,
          2,
          context,
        );
        assert.deepEqual(
          await resolvedView(batch, "accept"),
          expected,
          context,
        );
        assert.deepEqual(
          await resolvedView(batch, "reject"),
          original,
          context,
        );

        const sequential = await splitFinal(fixture, "tracked-changes");
        const operations = batchOf({
          reviewer: sequential,
          edit,
          deletionFirst,
        });
        for (const group of sequentialGroups({
          applied: operations,
          preRows: rowsOf(await resolved(sequential, "accept")),
          liveRows: rowsOf(sequential),
          mode: "tracked-changes",
        })) {
          const result = apply(
            sequential,
            "tracked-changes",
            group.map(({ id }) => {
              const operation = operations.find(
                (candidate) => candidate.id === id,
              );
              assert.ok(operation, `${context}: ${id} is the batch's`);
              return operation;
            }),
          );
          assert.deepEqual(result.skipped, [], context);
        }
        assert.deepEqual(
          await resolvedView(sequential, "accept"),
          expected,
          context,
        );
      }
    }
  }
});
