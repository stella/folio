/**
 * A comment a suggestion proposes is pending with it: listed as part of the
 * suggestion, left out of saves, carried by the host's pending-suggestion
 * export into a fresh session, removed on reject and kept on accept.
 */

import { describe, expect, test } from "bun:test";

import { createDocx } from "../docx/rezip";
import { paragraph } from "../docx/server/build";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";
import { createFolioAITextRangeHandle } from "./snapshot";
import type { FolioAIEditOperation } from "./types";

const FIRST_ID = "10000001";
const SECOND_ID = "10000002";
const SECOND_TEXT = "Second clause stays.";

const documentBytes = (): Promise<ArrayBuffer> => {
  const document = createEmptyDocument();
  document.package.document.content = [
    { ...paragraph("First clause remains."), paraId: FIRST_ID },
    { ...paragraph(SECOND_TEXT), paraId: SECOND_ID },
  ];
  return createDocx(document);
};

const range = () => {
  const handle = createFolioAITextRangeHandle({
    blockId: SECOND_ID,
    text: SECOND_TEXT,
    startOffset: 0,
    endOffset: 6,
  });
  if (!handle) throw new Error("range handle");
  return handle;
};

const OPERATIONS = {
  commentOnRange: (): FolioAIEditOperation => ({
    id: "proposal",
    type: "commentOnRange",
    range: range(),
    comment: { text: "Check this word." },
  }),
  commentOnBlock: (): FolioAIEditOperation => ({
    id: "proposal",
    type: "commentOnBlock",
    blockId: SECOND_ID,
    comment: { text: "Check this clause." },
  }),
  "replaceInBlock with a comment": (): FolioAIEditOperation => ({
    id: "proposal",
    type: "replaceInBlock",
    blockId: SECOND_ID,
    find: "Second",
    replace: "Third",
    comment: { text: "Check this edit." },
  }),
} as const;

const commentTexts = (reviewer: FolioDocxReviewer) =>
  reviewer.getComments().map(({ text }) => text);

const savedComments = async (reviewer: FolioDocxReviewer) =>
  commentTexts(await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer()));

const staged = async (operation: FolioAIEditOperation) => {
  const reviewer = await FolioDocxReviewer.fromBuffer(await documentBytes());
  const result = reviewer.applyOperations([operation], {
    undefinedReferences: "refuse",
    mode: "suggested",
  });
  expect(result.skipped).toEqual([]);
  expect(result.applied.at(0)?.suggestionId).toBe("proposal");
  return reviewer;
};

/** Export, save, reopen and restore: the host's round trip. */
const restored = async (reviewer: FolioDocxReviewer) => {
  const records = JSON.parse(
    JSON.stringify(reviewer.exportPendingSuggestions()),
  );
  expect(records).toHaveLength(1);
  expect(records.at(0)?.commentId).toBeDefined();
  const fresh = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
  expect(fresh.loadPendingSuggestions(records)).toEqual([
    { status: "restaged", suggestionId: "proposal" },
  ]);
  return fresh;
};

describe("suggested comments in the default engine", () => {
  for (const [name, operation] of Object.entries(OPERATIONS)) {
    test(`${name}: pending, saved without it, restored, rejected and accepted`, async () => {
      const reviewer = await staged(operation());
      const pending = (session: FolioDocxReviewer) =>
        session
          .exportPendingSuggestions()
          .map(({ suggestionId }) => suggestionId);
      expect(pending(reviewer)).toEqual(["proposal"]);
      // Pending: the save leaves the proposed comment out.
      expect(await savedComments(reviewer)).toEqual([]);

      const rejecting = await restored(reviewer);
      expect(pending(rejecting)).toEqual(["proposal"]);
      expect(rejecting.rejectSuggestion("proposal")).toBe(true);
      expect(pending(rejecting)).toEqual([]);
      expect(await savedComments(rejecting)).toEqual([]);
      expect(commentTexts(rejecting)).toEqual([]);

      const accepting = await restored(reviewer);
      expect(accepting.acceptSuggestion("proposal")).toBe(true);
      expect(pending(accepting)).toEqual([]);
      const proposed = operation();
      expect(await savedComments(accepting)).toEqual([
        "comment" in proposed ? proposed.comment?.text : undefined,
      ]);
    });
  }
});

describe("a proposed comment through a review of ordinary changes", () => {
  const reviewed = async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await documentBytes());
    reviewer.applyOperations(
      [
        {
          id: "ordinary",
          type: "replaceInBlock",
          blockId: FIRST_ID,
          find: "remains",
          replace: "stays",
        },
      ],
      { undefinedReferences: "refuse", mode: "tracked-changes" },
    );
    reviewer.applyOperations([OPERATIONS.commentOnRange()], {
      undefinedReferences: "refuse",
      mode: "suggested",
    });
    return reviewer;
  };
  const pending = (reviewer: FolioDocxReviewer) =>
    reviewer.exportPendingSuggestions().map(({ suggestionId }) => suggestionId);

  test("acceptAll keeps the proposal pending and its comment out of the file", async () => {
    const reviewer = await reviewed();
    expect(reviewer.acceptAll().resolved).toBeGreaterThan(0);
    expect(pending(reviewer)).toEqual(["proposal"]);
    expect(await savedComments(reviewer)).toEqual([]);
    expect(reviewer.acceptSuggestion("proposal")).toBe(true);
    expect(await savedComments(reviewer)).toEqual(["Check this word."]);
  });

  test("rejectAll removes the proposal with every other suggestion", async () => {
    const reviewer = await reviewed();
    expect(reviewer.rejectAll().resolved).toBeGreaterThan(0);
    expect(pending(reviewer)).toEqual([]);
    expect(await savedComments(reviewer)).toEqual([]);
    expect(commentTexts(reviewer)).toEqual([]);
  });
});
