/**
 * A comment on a range of a paragraph's text, as the reader lists it, quotes
 * exactly that range, and the oracle predicts it, beside every kind of
 * pending join: of an original paragraph, of one a reviewer inserted or
 * split off, or of one with pending inserted or deleted words, joined by
 * the reviewer who made those changes or by another one.
 */

import assert from "node:assert/strict";
import { test } from "node:test";

import { fromMarkdown } from "@stll/folio-core/markdown";
import {
  createFolioAITextRangeHandle,
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
} from "@stll/folio-core/server";

import { openReviewer, packDocument } from "../support/documents.ts";
import { SECOND_REVIEWER } from "../support/fuzz.ts";
import {
  assertRequestedOutcome,
  capture,
  expectOperation,
  modelOf,
  rowsOf,
} from "../support/oracle.ts";

type Reviewer = Awaited<ReturnType<typeof openReviewer>>;
type Operation = Parameters<
  Reviewer["applyDocumentOperations"]
>[0]["operations"][number];

const applyChecked = async (reviewer: Reviewer, operation: Operation) => {
  const pre = await capture(reviewer, "tracked-changes");
  const receipt = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      operations: [operation],
    },
    { undefinedReferences: "refuse" },
  );
  assert.equal(receipt.applied.length, 1, JSON.stringify(receipt));
  const model = modelOf(pre.rows, pre.liveRows, pre.joinSources);
  model.mode = pre.mode;
  expectOperation(model, operation);
  assert.deepEqual(model.unmodelled, [], operation.id);
  await assertRequestedOutcome(
    reviewer,
    pre,
    { applied: [operation] },
    operation.id,
  );
};

const TEXTS = ["Alpha clause words.", "Beta clause words.", "Gamma clause."];

/**
 * Pending changes the first reviewer makes before anyone joins, and how many
 * paragraphs the reader lists after them.
 */
const SETUPS: Record<
  string,
  { paragraphs: number; operations: (ids: readonly string[]) => Operation[] }
> = {
  "no other change": { paragraphs: 3, operations: () => [] },
  "an inserted paragraph": {
    paragraphs: 4,
    operations: ([first]) => [
      {
        id: "insert",
        type: "insertAfterBlock",
        blockId: String(first),
        text: "Added paragraph words.",
      },
    ],
  },
  "a split": {
    paragraphs: 4,
    operations: ([, second]) => [
      { id: "split", type: "splitBlock", blockId: String(second), offset: 5 },
    ],
  },
  "inserted words": {
    paragraphs: 3,
    operations: ([first]) => [
      {
        id: "insert-words",
        type: "replaceInBlock",
        blockId: String(first),
        find: "clause",
        replace: "clause added",
      },
    ],
  },
  "deleted words": {
    paragraphs: 3,
    operations: ([, second]) => [
      {
        id: "delete-words",
        type: "replaceInBlock",
        blockId: String(second),
        find: "clause ",
        replace: "",
      },
    ],
  },
};

/** Ranges at the start, the middle and the end of `text`, and all of it. */
const rangesOf = ({ length }: string): [number, number][] => [
  [0, length],
  [0, 3],
  [Math.floor(length / 2) - 1, Math.floor(length / 2) + 2],
  [length - 3, length],
];

const prepared = async (
  setup: (ids: readonly string[]) => Operation[],
  joiner: "same" | "other",
) => {
  let reviewer = await openReviewer(
    await packDocument(fromMarkdown(TEXTS.join("\n\n"))),
  );
  for (const operation of setup(rowsOf(reviewer).map(({ id }) => id)))
    await applyChecked(reviewer, operation);
  if (joiner === "other")
    reviewer = await openReviewer(
      new Uint8Array(await reviewer.toBuffer()),
      SECOND_REVIEWER,
    );
  return reviewer;
};

for (const [setup, { paragraphs, operations }] of Object.entries(SETUPS))
  for (const joiner of ["same", "other"] as const)
    for (let at = 0; at + 1 < paragraphs; at++)
      test(`${setup}, ${joiner} reviewer joins paragraph ${at + 1} on, then comments on ranges`, async () => {
        const reviewer = await prepared(operations, joiner);
        const live = rowsOf(reviewer);
        assert.equal(live.length, paragraphs);
        const source = live[at];
        const destination = live[at + 1];
        assert.ok(source && destination);
        await applyChecked(reviewer, {
          id: "join",
          type: "mergeBlockWithNext",
          blockId: source.id,
          separator: " ",
        });
        let count = 0;
        for (const id of [source.id, destination.id]) {
          // The reviewer's own inserted paragraph joins at once.
          const block = rowsOf(reviewer).find((row) => row.id === id);
          if (block === undefined) continue;
          for (const [startOffset, endOffset] of rangesOf(block.text)) {
            const range = createFolioAITextRangeHandle({
              blockId: id,
              text: block.text,
              startOffset,
              endOffset,
            });
            assert.ok(range);
            count += 1;
            const text = `Comment ${count}.`;
            await applyChecked(reviewer, {
              id: `comment-${count}`,
              type: "commentOnRange",
              range,
              comment: { text },
            });
            // The oracle checks the quote once accepted; the live one also
            // holds the words a range spans that are pending deletion.
            if (setup !== "deleted words")
              assert.equal(
                reviewer.getComments().find((added) => added.text === text)
                  ?.anchoredText,
                block.text.slice(startOffset, endOffset),
                `${id} ${startOffset}-${endOffset}`,
              );
          }
        }
        assert.ok(count > 0);
      });
