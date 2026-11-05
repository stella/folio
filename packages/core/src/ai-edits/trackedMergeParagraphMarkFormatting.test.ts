/**
 * A tracked merge hands the first paragraph's properties to the second, whose
 * words then read in that paragraph's cascade. A paragraph mark's run
 * properties lend its body runs only what a reopen reads from them: read with
 * the font as well, the live words took a look the save could not state, and
 * the same words read otherwise after the save.
 */

import { expect, test } from "bun:test";

import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import type { Paragraph } from "../types/document";
import { FolioDocxReviewer } from "./headless";

const buildReviewer = async () => {
  const document = fromMarkdown(
    "# Heading\n\nFirst paragraph.\n\nThe last one.",
  );
  const first = document.package.document.content.at(1) as Paragraph;
  first.formatting = {
    ...first.formatting,
    runProperties: {
      fontFamily: { ascii: "Mono Face", hAnsi: "Mono Face", cs: "Mono Face" },
      fontSize: 40,
      fontSizeCs: 40,
    },
  };
  return FolioDocxReviewer.fromBuffer(await createDocx(document), {
    author: "Test",
  });
};

const previewRuns = (reviewer: FolioDocxReviewer) =>
  reviewer.getContent().map((block) =>
    "previewRuns" in block
      ? (block.previewRuns ?? []).map(
          ({ text, fontSizePt, fontFamily, directFormatting }) => ({
            text,
            fontSizePt,
            fontFamily,
            directFormatting,
          }),
        )
      : [],
  );

test("the words a tracked merge hands a paragraph mark's cascade read the same after a save", async () => {
  const reviewer = await buildReviewer();
  const first = reviewer.getContent().at(1);
  if (!first) throw new Error("the first paragraph must exist");
  for (const operation of [
    {
      id: "insert",
      type: "insertAfterBlock",
      blockId: first.id,
      text: "Inserted heading words.",
      styleId: "Heading2",
    },
    {
      id: "merge",
      type: "mergeBlockWithNext",
      blockId: first.id,
      separator: " ",
    },
  ] as const) {
    const receipt = reviewer.applyDocumentOperations(
      {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "tracked-changes",
        operations: [operation],
      },
      { undefinedReferences: "refuse" },
    );
    expect(receipt.applied).toHaveLength(1);
  }
  const reopened = await FolioDocxReviewer.fromBuffer(
    await reviewer.toBuffer(),
    { author: "Test" },
  );
  expect(previewRuns(reopened)).toEqual(previewRuns(reviewer));
});
