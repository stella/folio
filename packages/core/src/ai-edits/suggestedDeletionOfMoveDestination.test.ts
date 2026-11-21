/**
 * A suggested deletion of an imported move's destination paragraph, then
 * Reject All: rejecting the move takes the destination's words and joins its
 * mark into the next paragraph, so nothing of the suggestion's paragraph
 * stands in the survivor and the suggestion is placed, not shifted.
 */
import { expect, test } from "bun:test";

import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { createDocx } from "../docx/rezip";
import type { BlockContent, Document, Paragraph } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";

const info = { author: "Earlier Reviewer", date: "2026-01-01T00:00:00Z" };
const run = (text: string) => ({
  type: "run" as const,
  content: [{ type: "text" as const, text }],
});
const plain = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [run(text)],
});

const DESTINATION = "00000004";
const ORIGINAL = [
  "Opening clause.",
  "Moved clause.",
  "Middle clause.",
  "Closing clause.",
];

const documentWith = (standing: BlockContent[]): Document => {
  const base = createEmptyDocument();
  return {
    ...base,
    package: {
      ...base.package,
      document: {
        ...base.package.document,
        content: [
          plain("00000001", "Opening clause."),
          {
            type: "paragraph",
            paraId: "00000002",
            content: [
              {
                type: "moveFrom",
                info: { ...info, id: 2 },
                content: [run("Moved clause.")],
              },
            ],
            pPrMark: { kind: "moveFrom", info: { ...info, id: 1 } },
          },
          plain("00000003", "Middle clause."),
          {
            type: "paragraph",
            paraId: DESTINATION,
            content: [
              { type: "bookmarkStart", id: 11, name: "MovedClause" },
              {
                type: "moveTo",
                info: { ...info, id: 4 },
                content: [run("Moved clause.")],
              },
            ],
            pPrMark: { kind: "moveTo", info: { ...info, id: 3 } },
          },
          ...standing,
          plain("00000005", "Closing clause."),
        ],
      },
    },
  };
};

for (const [name, standing] of [
  ["the next paragraph", []],
  [
    "a bookmark end before the next paragraph",
    [{ type: "bookmarkEnd", id: 11 }],
  ],
] as const)
  test(`canonical: Reject All after a suggested deletion of a move destination before ${name} gives the moves back`, async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(
      await createDocx(documentWith([...standing])),
      { session: "canonical", author: "Reviewer" },
    );
    const result = reviewer.applyDocumentOperations(
      {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "suggested",
        operations: [{ id: "op-1", type: "deleteBlock", blockId: DESTINATION }],
      },
      { undefinedReferences: "refuse" },
    );
    expect(result.skipped).toEqual([]);
    reviewer.rejectAll();
    expect(reviewer.snapshot().blocks.map(({ text }) => text)).toEqual(
      ORIGINAL,
    );
    expect(reviewer.getChanges()).toEqual([]);
  });
