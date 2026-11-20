/**
 * A comment or formatting over a range that holds whole note references
 * applies on the canonical session and keeps each reference in place; only a
 * range that cuts into one is refused.
 */
import { expect, test } from "bun:test";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { hashFolioAIBlockText } from "../ai-edits/snapshot";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import type { FolioDocumentOperation } from "../document-operations";
import { createDocx } from "../docx/rezip";
import { paragraph, run } from "../docx/server/build";
import { fromMarkdown } from "../markdown/fromMarkdown";

const open = async () => {
  const document = fromMarkdown("Opening words.");
  document.package.footnotes = [
    { type: "footnote", id: 1, content: [paragraph("As defined.")] },
  ];
  document.package.document.content.push(
    paragraph([
      run("Goods are defined."),
      { type: "run", content: [{ type: "footnoteRef", id: 1 }] },
    ]),
  );
  return FolioDocxReviewer.fromBuffer(await createDocx(document), {
    author: "Editor",
    session: "canonical",
  });
};

const apply = (
  reviewer: FolioDocxReviewer,
  operation: FolioDocumentOperation,
) =>
  reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      operations: [operation],
    },
    { undefinedReferences: "refuse" },
  );

const noteBlock = (reviewer: FolioDocxReviewer) => {
  const block = reviewer
    .snapshot()
    .blocks.find(({ text }) => text.startsWith("Goods are defined."));
  if (block === undefined) throw new Error("the fixture has no note block");
  return block;
};

test("a comment quoting a note reference applies and keeps it", async () => {
  const reviewer = await open();
  const block = noteBlock(reviewer);
  const result = apply(reviewer, {
    id: "comment",
    type: "commentOnBlock",
    blockId: block.id,
    quote: "[^1]",
    comment: { text: "Check the note." },
  });
  expect(result.skipped).toEqual([]);
  expect(reviewer.getComments().map(({ text }) => text)).toEqual([
    "Check the note.",
  ]);
  expect(noteBlock(reviewer).text).toBe(block.text);
});

test("formatting over words and a note reference applies and keeps it", async () => {
  const reviewer = await open();
  const block = noteBlock(reviewer);
  const startOffset = block.text.indexOf("defined.");
  const selected = block.text.slice(startOffset);
  const result = apply(reviewer, {
    id: "italic",
    type: "formatRange",
    range: {
      type: "textRange",
      story: "main",
      blockId: block.id,
      startOffset,
      endOffset: block.text.length,
      selectedTextHash: hashFolioAIBlockText(selected),
    },
    formatting: { italic: true },
  });
  expect(result.skipped).toEqual([]);
  expect(noteBlock(reviewer).text).toBe(block.text);
});

test("a range cutting into a note reference is still refused", async () => {
  const reviewer = await open();
  const block = noteBlock(reviewer);
  const endOffset = block.text.length - 1;
  const result = apply(reviewer, {
    id: "cut",
    type: "formatRange",
    range: {
      type: "textRange",
      story: "main",
      blockId: block.id,
      startOffset: 0,
      endOffset,
      selectedTextHash: hashFolioAIBlockText(block.text.slice(0, endOffset)),
    },
    formatting: { italic: true },
  });
  expect(result.skipped.map(({ reason }) => reason)).toEqual([
    "protectedReference",
  ]);
});
