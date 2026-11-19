import { panic } from "better-result";
import { expect, test } from "bun:test";

import { FolioDocxReviewer } from "../ai-edits/headless";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";

const open = async () =>
  FolioDocxReviewer.fromBuffer(
    (
      await prepareCanonicalDocxInput(
        await createDocx(
          fromMarkdown(
            [
              "Price Schedule",
              "The prices below apply.",
              "| Item | Price |\n|---|---|\n| Widget | 10 |",
              "Taxes are extra.",
            ].join("\n\n"),
          ),
        ),
      )
    )
      .unwrap()
      .slice().buffer,
    { session: "canonical" },
  );

const tracked = (
  reviewer: FolioDocxReviewer,
  operations: readonly FolioDocumentOperation[],
) => {
  const result = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      operations,
    },
    { undefinedReferences: "refuse" },
  );
  expect(result.skipped).toEqual([]);
  expect(result.applied).toHaveLength(operations.length);
};

test("a paragraph inserted before a table a pending join runs into goes before the join", async () => {
  const reviewer = await open();
  const id = (text: string) =>
    reviewer.getContent().find((block) => block.text === text)?.id ??
    panic(`The fixture lost "${text}".`);
  const heading = id("Price Schedule");
  const closing = id("The prices below apply.");
  const cell = id("Item");
  // The heading's break goes, then the paragraph it joins goes whole: once
  // accepted, the heading's words run into the table's first cell.
  tracked(reviewer, [
    {
      id: "merge",
      type: "mergeBlockWithNext",
      blockId: heading,
      separator: " ",
    },
  ]);
  tracked(reviewer, [{ id: "delete", type: "deleteBlock", blockId: closing }]);
  tracked(reviewer, [
    {
      id: "insert",
      type: "insertBeforeBlock",
      blockId: cell,
      text: "Clause words.",
    },
  ]);
  const accepted = reviewer.acceptAll();
  expect(accepted.skipped).toEqual([]);
  expect(reviewer.getContent().map(({ text }) => text)).toEqual([
    "Clause words.",
    "Price Schedule Item",
    "Price",
    "Widget",
    "10",
    "Taxes are extra.",
  ]);
});
