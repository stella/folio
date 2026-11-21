import { expect, test } from "bun:test";

import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { getParagraphText } from "../docx/paragraphParser";
import { createDocx } from "../docx/rezip";
import type { Document, Paragraph, SectionProperties } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";

const BOUNDARY = "00000002";
const ENDED_SECTION = { pageWidth: 10_000 } satisfies SectionProperties;

const paragraph = (
  paraId: string,
  text: string,
  sectionProperties?: SectionProperties,
): Paragraph => ({
  type: "paragraph",
  paraId,
  textId: paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
  ...(sectionProperties === undefined ? {} : { sectionProperties }),
});

const source = (): Document => {
  const document = createEmptyDocument();
  return {
    ...document,
    package: {
      ...document.package,
      document: {
        ...document.package.document,
        content: [
          paragraph("00000001", "Opening."),
          paragraph(BOUNDARY, "Boundary.", ENDED_SECTION),
          paragraph("00000003", "Following."),
        ],
        finalSectionProperties: { pageWidth: 14_000 },
      },
    },
  };
};

/** Each paragraph's text and the page width of the section it ends, if any. */
const sections = (reviewer: FolioDocxReviewer) =>
  reviewer.toDocument().package.document.content.flatMap((block) =>
    block.type === "paragraph"
      ? [
          {
            text: getParagraphText(block),
            pageWidth: block.sectionProperties?.pageWidth ?? null,
          },
        ]
      : [],
  );

const insertAfterBoundary = {
  id: "insert-after-boundary",
  type: "insertAfterBlock",
  blockId: BOUNDARY,
  text: "Inserted.",
} as const satisfies FolioDocumentOperation;
const deleteBoundary = {
  id: "delete-boundary",
  type: "deleteBlock",
  blockId: BOUNDARY,
} as const satisfies FolioDocumentOperation;

const cases = [
  {
    name: "an insertion after it",
    operations: [insertAfterBoundary],
    direct: [
      { text: "Opening.", pageWidth: null },
      { text: "Boundary.", pageWidth: 10_000 },
      { text: "Inserted.", pageWidth: null },
      { text: "Following.", pageWidth: null },
    ],
  },
  {
    // The inserted paragraph never takes the deleted paragraph's break, so
    // rejecting both restores it and accepting both deletes it, as a direct
    // deletion does.
    name: "its deletion and an insertion after it",
    operations: [deleteBoundary, insertAfterBoundary],
    direct: [
      { text: "Opening.", pageWidth: null },
      { text: "Inserted.", pageWidth: null },
      { text: "Following.", pageWidth: null },
    ],
  },
] as const;

const open = (
  bytes: Uint8Array | ArrayBuffer,
  session: "legacy" | "canonical",
) => FolioDocxReviewer.fromBuffer(bytes, { session, author: "Reviewer" });

for (const session of ["legacy", "canonical"] as const)
  for (const { name, operations, direct } of cases) {
    test(`${session}: ${name} keeps a section-ending paragraph's break on its own mark`, async () => {
      const bytes = await createDocx(source());
      const original = sections(await open(bytes, session));

      const directReviewer = await open(bytes, session);
      const directResult = directReviewer.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode: "direct",
          operations,
        },
        { undefinedReferences: "refuse" },
      );
      expect(directResult.applied).toHaveLength(operations.length);
      expect(sections(directReviewer)).toEqual(direct);

      const tracked = await open(bytes, session);
      const trackedResult = tracked.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode: "tracked-changes",
          operations,
        },
        { undefinedReferences: "refuse" },
      );
      expect(trackedResult.applied).toHaveLength(operations.length);
      const saved = await tracked.toBuffer();

      const rejected = await open(saved, session);
      rejected.rejectAll();
      expect(sections(rejected)).toEqual(original);

      const accepted = await open(saved, session);
      accepted.acceptAll();
      expect(sections(accepted)).toEqual(direct);
    });
  }
