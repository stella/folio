/**
 * Deleting a paragraph that ends a section on the canonical session, through
 * a save: the deleted mark keeps its section break, so rejecting restores
 * every section and accepting joins the two around the break under the
 * properties of the one after it.
 */
import { expect, test } from "bun:test";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { createDocx } from "../docx/rezip";
import { paragraph } from "../docx/server/build";
import { createEmptyDocument } from "../utils/createDocument";

const open = async (bytes: ArrayBuffer) =>
  FolioDocxReviewer.fromBuffer(bytes, {
    author: "Editor",
    session: "canonical",
  });

const source = () => {
  const document = createEmptyDocument();
  document.package.document.content = [
    {
      ...paragraph("First section."),
      sectionProperties: { pageWidth: 10_000 },
    },
    {
      ...paragraph("Second section."),
      sectionProperties: { pageWidth: 12_000 },
    },
    paragraph("Final section."),
  ];
  return createDocx(document);
};

const pageWidths = (reviewer: FolioDocxReviewer) =>
  reviewer
    .toDocument()
    .package.document.content.flatMap((block) =>
      block.type === "paragraph" && block.sectionProperties !== undefined
        ? [block.sectionProperties.pageWidth]
        : [],
    );

const deletedBoundary = async () => {
  const reviewer = await open(await source());
  const boundary = reviewer
    .getContent()
    .find(({ text }) => text === "First section.");
  if (boundary === undefined) throw new Error("The boundary reads.");
  const result = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      operations: [{ id: "delete", type: "deleteBlock", blockId: boundary.id }],
    },
    { undefinedReferences: "refuse" },
  );
  expect(result.skipped).toEqual([]);
  return open(await reviewer.toBuffer());
};

test("a saved tracked deletion of a section's last paragraph keeps its break pending", async () => {
  const reopened = await deletedBoundary();
  expect(pageWidths(reopened)).toEqual([10_000, 12_000]);
  reopened.rejectAll();
  const rejected = await open(await reopened.toBuffer());
  expect(pageWidths(rejected)).toEqual([10_000, 12_000]);
  expect(rejected.getContent().map(({ text }) => text)).toEqual([
    "First section.",
    "Second section.",
    "Final section.",
  ]);
});

test("accepting it joins the section into the one after its break", async () => {
  const reopened = await deletedBoundary();
  reopened.acceptAll();
  const accepted = await open(await reopened.toBuffer());
  expect(pageWidths(accepted)).toEqual([12_000]);
  expect(accepted.getContent().map(({ text }) => text)).toEqual([
    "Second section.",
    "Final section.",
  ]);
});
