/**
 * A block put before a paragraph lands before the paragraph before it that
 * accepting removes whole, at the start of that gap, so a block put after the
 * removed paragraph later in the batch follows it. One the batch put after the
 * removed paragraph first stops that gap, as it does one at a time: the two
 * keep the batch's order.
 */
import { expect, test } from "bun:test";

import { FolioDocxReviewer } from "../ai-edits/headless";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";

const tracked = (operations: FolioDocumentOperation[]) => ({
  version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  mode: "tracked-changes" as const,
  operations,
});

const texts = (reviewer: FolioDocxReviewer) =>
  reviewer
    .snapshot()
    .blocks.map(({ text }) => text)
    .filter((text) => text !== "");

test.each([
  ["the paragraph first", ["paragraph", "table"]],
  ["the table first", ["table", "paragraph"]],
] as const)(
  "a block put before the paragraph after a removed one goes before it (%s)",
  async (_name, order) => {
    const reviewer = await FolioDocxReviewer.fromBuffer(
      await createDocx(
        fromMarkdown("Opening words.\n\nRemoved words.\n\nClosing words."),
      ),
      { author: "Editor", session: "canonical" },
    );
    const [, removed, closing] = reviewer.snapshot().blocks;
    if (removed === undefined || closing === undefined)
      throw new Error("the fixture's blocks are missing");
    expect(
      reviewer.applyDocumentOperations(
        tracked([{ id: "delete", type: "deleteBlock", blockId: removed.id }]),
        { undefinedReferences: "refuse" },
      ).skipped,
    ).toEqual([]);
    const operations: Record<string, FolioDocumentOperation> = {
      paragraph: {
        id: "paragraph",
        type: "insertBeforeBlock",
        blockId: closing.id,
        text: "Inserted words.",
      },
      table: {
        id: "table",
        type: "insertTable",
        blockId: removed.id,
        rows: [["Term", "Value"]],
      },
    };
    const result = reviewer.applyDocumentOperations(
      tracked(order.map((name) => operations[name]!)),
      { undefinedReferences: "refuse" },
    );
    expect(result.skipped).toEqual([]);
    const named =
      order[0] === "paragraph"
        ? ["Inserted words.", "Term", "Value"]
        : ["Term", "Value", "Inserted words."];
    expect(texts(reviewer)).toEqual([
      "Opening words.",
      ...named,
      "Closing words.",
    ]);
    reviewer.acceptAll();
    expect(texts(reviewer)).toEqual([
      "Opening words.",
      ...named,
      "Closing words.",
    ]);
  },
);
