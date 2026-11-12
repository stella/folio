/**
 * Deleting a run of whole paragraphs, one of which carries a comment, reads
 * the same in every mode: applied directly, as tracked changes accepted, and
 * as suggestions accepted. The comment's whole range goes with the words, and
 * so does its reference, so nothing is left of the commented paragraph to
 * hold an empty one in place: accepting leaves no blank paragraph where the
 * run was, whichever paragraph of the run the comment is on and whether the
 * run is deleted in one batch or one paragraph per batch.
 */

import { describe, expect, test } from "bun:test";

import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { FolioDocxReviewer } from "./headless";

type Mode = "direct" | "tracked-changes" | "suggested";

const KEPT_BEFORE = "Opening clause.";
const RUN = [
  "First deleted clause.",
  "Middle deleted clause.",
  "Last deleted clause.",
];
const KEPT_AFTER = "Closing clause.";

const open = async (): Promise<FolioDocxReviewer> => {
  const markdown = [
    `## ${KEPT_BEFORE}`,
    ...RUN,
    "- Listed clause.",
    KEPT_AFTER,
  ];
  const bytes = (
    await ensureParaIds(await createDocx(fromMarkdown(markdown.join("\n\n"))))
  ).docx;
  return FolioDocxReviewer.fromBuffer(bytes, { author: "Reviewer" });
};

const reopen = async (
  reviewer: FolioDocxReviewer,
): Promise<FolioDocxReviewer> =>
  FolioDocxReviewer.fromBuffer(await reviewer.toBuffer(), {
    author: "Reviewer",
  });

/** Apply `operations` in `mode`, accepting what it suggests as tracked changes. */
const applyBatch = (
  reviewer: FolioDocxReviewer,
  mode: Mode,
  operations: FolioDocumentOperation[],
) => {
  const result = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode,
      atomic: false,
      operations,
    },
    { undefinedReferences: "refuse" },
  );
  expect(result.skipped).toEqual([]);
  const suggestionIds = new Set(
    result.applied.flatMap(({ suggestionId }) => suggestionId ?? []),
  );
  for (const suggestionId of suggestionIds) {
    expect(reviewer.acceptSuggestion(suggestionId)).toBe(true);
  }
};

type Setup = {
  /** Which paragraph of the run carries the comment. */
  commented: 0 | 1 | 2;
  batching: "one batch" | "one per batch";
  /** The paragraph the deletions start from. */
  order: "first to last" | "last to first";
};

/** The blocks once every change is accepted, saved and reopened. */
const replay = async (setup: Setup, mode: Mode) => {
  const reviewer = await open();
  const blocks = reviewer.snapshot().blocks;
  const ids = RUN.map((text) => {
    const block = blocks.find((candidate) => candidate.text === text);
    if (!block) throw new Error(`No block reads "${text}"`);
    return block.id;
  });
  const commentedId = ids[setup.commented];
  if (commentedId === undefined) throw new Error("No commented block");
  applyBatch(reviewer, "direct", [
    {
      id: "comment",
      type: "commentOnBlock",
      blockId: commentedId,
      comment: { text: "Check this clause." },
    },
  ]);
  const deletions = ids.map((blockId, index): FolioDocumentOperation => ({
    id: `delete-${index}`,
    type: "deleteBlock",
    blockId,
  }));
  const ordered =
    setup.order === "first to last" ? deletions : deletions.toReversed();
  const batches =
    setup.batching === "one batch"
      ? [ordered]
      : ordered.map((deletion) => [deletion]);
  for (const operations of batches) applyBatch(reviewer, mode, operations);
  if (mode !== "direct") reviewer.acceptAll();
  const accepted = await reopen(reviewer);
  return {
    blocks: accepted.snapshot().blocks.map((block) => ({
      text: block.text,
      styleId: block.styleId,
      numbering: block.statedNumbering?.kind,
    })),
    comments: accepted.getComments().map((comment) => comment.anchoredText),
  };
};

const setups: Setup[] = ([0, 1, 2] as const).flatMap((commented) =>
  (["one batch", "one per batch"] as const).flatMap((batching) =>
    (["first to last", "last to first"] as const).map((order) => ({
      commented,
      batching,
      order,
    })),
  ),
);

describe("a run of whole-paragraph deletions with a commented paragraph", () => {
  for (const setup of setups) {
    const name = `comment on the ${["first", "middle", "last"][setup.commented]}, ${setup.batching}, ${setup.order}`;
    test(`leaves no paragraph behind in any mode (${name})`, async () => {
      const direct = await replay(setup, "direct");
      expect(direct.blocks.map(({ text }) => text)).toEqual([
        KEPT_BEFORE,
        "Listed clause.",
        KEPT_AFTER,
      ]);
      expect(await replay(setup, "tracked-changes")).toEqual(direct);
      expect(await replay(setup, "suggested")).toEqual(direct);
    });
  }
});
