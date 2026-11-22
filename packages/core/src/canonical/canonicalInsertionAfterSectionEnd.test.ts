/**
 * A paragraph put directly after one that ends a section, before a block that
 * is no paragraph, on the canonical session: it opens the section after the
 * break, which stays where it was, in the saved file too, and undoes.
 */
import { expect, test } from "bun:test";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { createDocx } from "../docx/rezip";
import type { BlockContent, Paragraph } from "../types/document";

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const FOLLOWERS: Record<string, BlockContent> = {
  "a content control": {
    type: "blockSdt",
    properties: { sdtType: "richText", id: 4 },
    content: [paragraph("00000003", "Held.")],
  },
  "a table": {
    type: "table",
    columnWidths: [2000],
    rows: [
      {
        type: "tableRow",
        cells: [
          { type: "tableCell", content: [paragraph("00000003", "Held.")] },
        ],
      },
    ],
  },
};

const open = (buffer: ArrayBuffer) =>
  FolioDocxReviewer.fromBuffer(buffer.slice(0), {
    author: "Editor",
    session: "canonical",
  });

/** Each section's page width and the words of its paragraphs outside blocks. */
const sections = (reviewer: FolioDocxReviewer) =>
  (reviewer.toDocument().package.document.sections ?? []).map(
    ({ properties, content }) => ({
      width: properties.pageWidth,
      blocks: content.map((block) =>
        block.type === "paragraph" ? block.paraId : block.type,
      ),
    }),
  );

test.each(
  Object.entries(FOLLOWERS).map(([kind, follower]) => ({ kind, follower })),
)("opens the section after the break, before $kind", async ({ follower }) => {
  const buffer = await createDocx({
    package: {
      document: {
        content: [
          {
            ...paragraph("00000001", "First."),
            sectionProperties: { pageWidth: 11_906 },
          },
          follower,
          paragraph("00000002", "Last."),
        ],
        finalSectionProperties: { pageWidth: 12_240 },
      },
    },
  });
  const reviewer = await open(buffer);
  const before = sections(reviewer);
  const { applied, skipped, undoHandle } = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "direct",
      operations: [
        {
          id: "op",
          type: "insertAfterBlock",
          blockId: "00000001",
          text: "New.",
        },
      ],
    },
    { undefinedReferences: "refuse" },
  );
  expect(skipped).toEqual([]);
  expect(applied.map(({ id }) => id)).toEqual(["op"]);
  const after = sections(reviewer);
  expect(after.map(({ width }) => width)).toEqual([11_906, 12_240]);
  expect(after.at(0)).toEqual(before.at(0));
  expect(after.at(1)?.blocks.slice(1)).toEqual(before.at(1)?.blocks ?? []);
  expect(reviewer.getContent().map(({ text }) => text)).toEqual([
    "First.",
    "New.",
    "Held.",
    "Last.",
  ]);
  const reopened = await open(await reviewer.toBuffer());
  expect(sections(reopened)).toEqual(after);
  if (undoHandle === null) throw new Error("The batch has an undo handle.");
  expect(reviewer.undoDocumentOperations(undoHandle)).toMatchObject({
    status: "undone",
  });
  expect(sections(reviewer)).toEqual(before);
});
