/**
 * A merge's separator stands between the words of the two paragraphs. When
 * every paragraph after the first is pending deletion, accepting leaves no
 * words after the separator, so the canonical session writes none, as the
 * default engine does; with words after it, the separator stays.
 */

import { expect, test } from "bun:test";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";

const open = async () => {
  const input = (
    await prepareCanonicalDocxInput(
      await createDocx(fromMarkdown("Opening words.\n\nClosing words.")),
    )
  )
    .unwrap()
    .slice().buffer;
  const reviewer = await FolioDocxReviewer.fromBuffer(input, {
    author: "Test",
    session: "canonical",
  });
  const [opening, closing] = reviewer.snapshot().blocks;
  if (opening === undefined || closing?.text !== "Closing words.")
    throw new Error("Missing fixture");
  return { reviewer, opening, closing };
};

const merge = (reviewer: FolioDocxReviewer, blockId: string) => {
  const result = reviewer.applyOperations(
    [{ id: "merge", type: "mergeBlockWithNext", blockId, separator: " " }],
    { mode: "tracked-changes" },
  );
  expect(result.skipped).toEqual([]);
};

const accepted = (reviewer: FolioDocxReviewer) => {
  reviewer.acceptAll();
  return reviewer.snapshot().blocks.map(({ text }) => text);
};

test("a merge into a paragraph of deleted words writes no separator", async () => {
  const { reviewer, opening, closing } = await open();
  expect(
    reviewer.applyOperations(
      [{ id: "delete", type: "deleteBlock", blockId: closing.id }],
      { mode: "tracked-changes" },
    ).skipped,
  ).toEqual([]);
  // Keep the closing paragraph's break: only its words are pending deletion.
  for (const change of reviewer.getChanges())
    if (change.type === "paragraphMarkDeleted")
      expect(reviewer.rejectChange(change)).toBe(true);
  merge(reviewer, opening.id);
  expect(accepted(reviewer)).toEqual(["Opening words."]);
});

test("a merge into a paragraph with words keeps its separator", async () => {
  const { reviewer, opening } = await open();
  merge(reviewer, opening.id);
  expect(accepted(reviewer)).toEqual(["Opening words. Closing words."]);
});
