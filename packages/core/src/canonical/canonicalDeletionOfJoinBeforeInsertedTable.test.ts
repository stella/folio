/**
 * A tracked deletion of the paragraph a pending join runs into, right before
 * a table the same author inserted after it: the table took over the
 * paragraph's break, so its own mark is that author's pending insertion. The
 * deletion takes the joined words and records itself on that mark, so
 * accepting everything removes the paragraph, as the batch that deletes it
 * and puts the table after it does, and as deleting the joined paragraph
 * directly does. No empty paragraph stays before the table.
 */
import { describe, expect, test } from "bun:test";

import { FolioDocxReviewer } from "../ai-edits/headless";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";

const IDS = ["0000000A", "0000000B", "0000000C"] as const;

const open = async (markdown: string) => {
  const document = fromMarkdown(markdown);
  document.package.document.content.forEach((block, index) => {
    if (block.type === "paragraph") block.paraId = IDS[index];
  });
  const prepared = await prepareCanonicalDocxInput(await createDocx(document));
  if (prepared.isErr()) throw prepared.error;
  return FolioDocxReviewer.fromBuffer(prepared.value.slice().buffer, {
    author: "Editor",
    session: "canonical",
  });
};

const apply = (
  reviewer: FolioDocxReviewer,
  operations: FolioDocumentOperation[],
) =>
  reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      operations,
    },
    { undefinedReferences: "refuse" },
  ).skipped;

const MERGE: FolioDocumentOperation = {
  id: "merge",
  type: "mergeBlockWithNext",
  blockId: "0000000A",
  separator: " ",
};
const DELETE: FolioDocumentOperation = {
  id: "delete",
  type: "deleteBlock",
  blockId: "0000000B",
};
const TABLE: FolioDocumentOperation = {
  id: "table",
  type: "insertTable",
  blockId: "0000000B",
  rows: [
    ["Term", "Value"],
    ["Period", "12 months"],
  ],
};

const accepted = (reviewer: FolioDocxReviewer) => {
  expect(reviewer.acceptAll().skipped).toEqual([]);
  return reviewer
    .getContent()
    .map(({ text, table }) => (table === undefined ? text : `[${text}]`));
};

const DOCUMENTS = {
  "ending the story": "One\n\nTwo",
  "before another paragraph": "One\n\nTwo\n\nThree",
} as const;

describe("a tracked deletion of a joined paragraph before its author's table", () => {
  for (const [where, markdown] of Object.entries(DOCUMENTS))
    test(`goes whole once accepted, as the batch does (${where})`, async () => {
      const together = await open(markdown);
      expect(apply(together, [MERGE])).toEqual([]);
      expect(apply(together, [DELETE, TABLE])).toEqual([]);

      const apart = await open(markdown);
      expect(apply(apart, [MERGE])).toEqual([]);
      expect(apply(apart, [TABLE])).toEqual([]);
      expect(apply(apart, [DELETE])).toEqual([]);

      const result = accepted(apart);
      expect(result).toEqual(accepted(together));
      expect(result.slice(0, 4)).toEqual([
        "[Term]",
        "[Value]",
        "[Period]",
        "[12 months]",
      ]);
    });
});
