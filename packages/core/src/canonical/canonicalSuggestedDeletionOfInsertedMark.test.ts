/**
 * A suggestion deleting a paragraph whose break another author's tracked
 * insertion added: the break holds the insertion and, beside it, the
 * suggestion's deletion. That deletion is the suggestion's, so it stays
 * staged through accept-all, goes with the suggestion's rejection, and is
 * never saved.
 */
import { expect, test } from "bun:test";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { createDocx } from "../docx/rezip";
import { paragraph } from "../docx/server/build";
import { fromMarkdown } from "../markdown/fromMarkdown";

const INSERTED = { id: 1, author: "Alice", date: "2026-01-01T00:00:00Z" };

const open = async () => {
  const document = fromMarkdown("Opening words.");
  document.package.document.content.push(
    { ...paragraph("Added clause."), pPrMark: { kind: "ins", info: INSERTED } },
    paragraph("Closing words."),
  );
  const reviewer = await FolioDocxReviewer.fromBuffer(
    await createDocx(document),
    { author: "Bob", session: "canonical" },
  );
  const added = reviewer
    .getContent()
    .find(({ text }) => text === "Added clause.");
  if (added === undefined) throw new Error("The added paragraph reads.");
  const result = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "suggested",
      operations: [{ id: "delete", type: "deleteBlock", blockId: added.id }],
    },
    { undefinedReferences: "refuse" },
  );
  expect(result.skipped).toEqual([]);
  return reviewer;
};

const texts = (reviewer: FolioDocxReviewer) =>
  reviewer.getContent().map(({ text }) => text);

const SOURCE_TEXTS = ["Opening words.", "Added clause.", "Closing words."];

const SUGGESTION = "delete";

test("accept-all leaves the suggestion's deletion of an inserted break staged", async () => {
  const reviewer = await open();
  // The suggestion deletes the paragraph's words; its break still stands.
  const suggested = ["Opening words.", "", "Closing words."];
  expect(texts(reviewer)).toEqual(suggested);
  reviewer.acceptAll();
  expect(texts(reviewer)).toEqual(suggested);
  expect(reviewer.rejectSuggestion(SUGGESTION)).toBe(true);
  expect(texts(reviewer)).toEqual(SOURCE_TEXTS);
  expect(reviewer.getChanges()).toEqual([]);
});

test("rejecting the suggestion leaves the inserted break pending", async () => {
  const reviewer = await open();
  expect(reviewer.rejectSuggestion(SUGGESTION)).toBe(true);
  expect(texts(reviewer)).toEqual(SOURCE_TEXTS);
  expect(reviewer.getChanges().map(({ id }) => id)).toEqual([INSERTED.id]);
});

test("a save keeps the inserted break without the suggestion", async () => {
  const reviewer = await open();
  const reopened = await FolioDocxReviewer.fromBuffer(
    await reviewer.toBuffer(),
    { author: "Bob", session: "canonical" },
  );
  expect(texts(reopened)).toEqual(SOURCE_TEXTS);
  expect(reopened.getChanges().map(({ id }) => id)).toEqual([INSERTED.id]);
});
