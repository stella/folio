import { panic } from "better-result";
import { expect, test } from "bun:test";

import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
  type FolioDocumentOperationBatch,
} from "../document-operations";
import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { table } from "../docx/server/build";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { FolioDocxReviewer } from "./headless";

const openFixture = async (): Promise<FolioDocxReviewer> => {
  const document = fromMarkdown(
    "Before paragraph.\n\nTarget paragraph.\n\nAfter paragraph.",
  );
  document.package.document.content.splice(
    2,
    0,
    table({ header: ["Item", "Value"], rows: [["Term", "Definition"]] }),
  );
  const { docx } = await ensureParaIds(
    new Uint8Array(await createDocx(document)),
  );
  const buffer = new Uint8Array(docx).buffer;
  return FolioDocxReviewer.fromBuffer(buffer, { author: "Reviewer" });
};

const read = (reviewer: FolioDocxReviewer): string[] =>
  reviewer.getContent().map(({ kind, text }) => `${kind}: ${text}`);

const apply = (
  reviewer: FolioDocxReviewer,
  operation: FolioDocumentOperation,
): number => {
  const batch = {
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    mode: "tracked-changes",
    operations: [operation],
  } satisfies FolioDocumentOperationBatch;
  return reviewer.applyDocumentOperations(batch, {
    undefinedReferences: "refuse",
  }).applied.length;
};

test("reject removes adjacent insertions and a comment without adding a paragraph before a table", async () => {
  const reviewer = await openFixture();
  const original = read(reviewer);
  const target = reviewer
    .getContent()
    .find(({ text }) => text === "Target paragraph.");
  if (!target) return panic("Missing target paragraph");

  const firstBefore = new Set(reviewer.getContent().map(({ id }) => id));
  expect(
    apply(reviewer, {
      id: "insert-first",
      type: "insertAfterBlock",
      blockId: target.id,
      text: "First inserted clause.",
    }),
  ).toBe(1);
  const first = reviewer.getContent().find(({ id }) => !firstBefore.has(id));
  if (!first) return panic("Missing first inserted paragraph");

  const secondBefore = new Set(reviewer.getContent().map(({ id }) => id));
  expect(
    apply(reviewer, {
      id: "insert-second",
      type: "insertAfterBlock",
      blockId: target.id,
      text: "Second inserted clause.",
    }),
  ).toBe(1);
  const second = reviewer.getContent().find(({ id }) => !secondBefore.has(id));
  if (!second) return panic("Missing second inserted paragraph");
  expect(first.id).not.toBe(second.id);

  expect(
    apply(reviewer, {
      id: "comment-second",
      type: "commentOnBlock",
      blockId: second.id,
      quote: "Second",
      comment: { text: "Review this clause." },
    }),
  ).toBe(1);

  reviewer.rejectAll();

  expect(
    reviewer.getComments().map(({ anchoredText }) => anchoredText),
  ).toEqual([]);
  expect(read(reviewer)).toEqual(original);

  const reloaded = await FolioDocxReviewer.fromBuffer(
    await reviewer.toBuffer(),
    {
      author: "Reviewer",
    },
  );
  expect(read(reloaded)).toEqual(original);
  expect(
    reloaded.getComments().map(({ anchoredText }) => anchoredText),
  ).toEqual([]);
});

test("reject removes adjacent insertions before a table without comments", async () => {
  const reviewer = await openFixture();
  const original = read(reviewer);
  const target = reviewer
    .getContent()
    .find(({ text }) => text === "Target paragraph.");
  if (!target) return panic("Missing target paragraph");

  expect(
    apply(reviewer, {
      id: "insert-first",
      type: "insertAfterBlock",
      blockId: target.id,
      text: "First inserted clause.",
    }),
  ).toBe(1);
  expect(
    apply(reviewer, {
      id: "insert-second",
      type: "insertAfterBlock",
      blockId: target.id,
      text: "Second inserted clause.",
    }),
  ).toBe(1);

  reviewer.rejectAll();

  expect(read(reviewer)).toEqual(original);
});
