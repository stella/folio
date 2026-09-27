import { expect, test } from "bun:test";

import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { FolioDocxReviewer } from "./headless";

const open = async (markdown: string) => {
  const bytes = (await ensureParaIds(await createDocx(fromMarkdown(markdown)))).docx;
  return FolioDocxReviewer.fromBuffer(bytes, { author: "Test" });
};

const reopen = async (reviewer: FolioDocxReviewer) =>
  FolioDocxReviewer.fromBuffer(await reviewer.toBuffer(), { author: "Test" });

const blockId = (reviewer: FolioDocxReviewer, text: string) => {
  const block = reviewer.getContent().find((candidate) => candidate.text === text);
  if (!block) throw new Error(`Missing paragraph: ${text}`);
  return block.id;
};

const previewOf = (reviewer: FolioDocxReviewer, text: string) =>
  reviewer.getContent().find((block) => block.text === text)?.previewRuns;

test.each(["direct", "tracked-changes"] as const)(
  "inserted table cells resolve default run styles before save (%s)",
  async (mode) => {
    const reviewer = await open("# Heading\n\nBody clause.");
    reviewer.applyDocumentOperations({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode,
      operations: [
        {
          id: "table",
          type: "insertTable",
          blockId: blockId(reviewer, "Body clause."),
          rows: [
            ["Term", "Value"],
            ["Period", "12 months"],
          ],
        },
      ],
    });
    const saved = await reopen(reviewer);
    for (const cellText of ["Term", "Value", "Period", "12 months"]) {
      expect(previewOf(reviewer, cellText)).toEqual(previewOf(saved, cellText));
      expect(previewOf(reviewer, cellText)).toEqual([
        { text: cellText, fontSizePt: 11, fontFamily: "Arial" },
      ]);
    }
  },
);

test("a tracked merge separator uses its owning paragraph's run style", async () => {
  const reviewer = await open(
    "# Removed heading\n\nFirst body.\n\nSecond body.\n\n**Bold tail**\n\n# Following heading",
  );
  const heading = blockId(reviewer, "Removed heading");
  const first = blockId(reviewer, "First body.");
  const bold = blockId(reviewer, "Bold tail");
  reviewer.applyDocumentOperations({
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    mode: "tracked-changes",
    operations: [{ id: "delete-heading", type: "deleteBlock", blockId: heading }],
  });
  reviewer.acceptAll();
  reviewer.applyDocumentOperations({
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    mode: "tracked-changes",
    operations: [{ id: "merge-first", type: "mergeBlockWithNext", blockId: first, separator: " " }],
  });
  reviewer.acceptAll();
  const secondMerge = reviewer.applyDocumentOperations({
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    mode: "tracked-changes",
    operations: [
      { id: "merge-second", type: "mergeBlockWithNext", blockId: first, separator: " " },
    ],
  });
  expect(secondMerge.applied).toHaveLength(1);
  const merged = reviewer.getContent().find((block) => block.id === first);
  expect(merged?.previewRuns).toBeDefined();
  expect(merged?.previewRuns).toEqual(
    (await reopen(reviewer)).getContent().find((block) => block.id === first)?.previewRuns,
  );
  reviewer.applyDocumentOperations({
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    mode: "tracked-changes",
    operations: [{ id: "merge-bold", type: "mergeBlockWithNext", blockId: bold, separator: " " }],
  });
  const boldPreview = reviewer.getContent().find((block) => block.id === bold)?.previewRuns;
  expect(boldPreview?.at(-1)).toEqual(expect.objectContaining({ bold: true, fontSizePt: 11 }));
  expect(boldPreview?.at(-1)?.text.endsWith(" ")).toBe(true);
  expect(boldPreview).toEqual(
    (await reopen(reviewer)).getContent().find((block) => block.id === bold)?.previewRuns,
  );
});

test("a direct merge after a tracked replacement reads the same run formatting after save", async () => {
  const reviewer = await open("# Price Schedule\n\nThe prices below apply.\n\nTaxes are extra.");
  const heading = blockId(reviewer, "Price Schedule");
  reviewer.applyDocumentOperations({
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    mode: "tracked-changes",
    operations: [
      { id: "replace-heading", type: "replaceBlock", blockId: heading, text: "Delivery notice." },
    ],
  });
  const merged = reviewer.applyDocumentOperations({
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    mode: "direct",
    operations: [{ id: "merge", type: "mergeBlockWithNext", blockId: heading, separator: " " }],
  });
  expect(merged.applied).toHaveLength(1);
  expect(reviewer.getContent()).toEqual((await reopen(reviewer)).getContent());
});

test("a later tracked merge does not borrow formatting from a resolved heading", async () => {
  const reviewer = await open(
    "# Service Agreement\n\nThis agreement is made between the parties named below.\n\nThe Supplier delivers the goods on time and in good order.\n\nThe Buyer pays each invoice within thirty days.\n\nSigned in two copies.",
  );
  const [heading, clause, supplier, , closing] = reviewer.getContent().map((block) => block.id);
  if (!heading || !clause || !supplier || !closing) throw new Error("Missing fixture paragraphs");
  const apply = (
    mode: "suggested" | "tracked-changes",
    operations: Parameters<FolioDocxReviewer["applyDocumentOperations"]>[0]["operations"],
  ) =>
    reviewer.applyDocumentOperations({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode,
      operations,
    });
  apply("suggested", [
    { id: "suggest-delete", type: "deleteBlock", blockId: heading },
    {
      id: "suggest-replace",
      type: "replaceInBlock",
      blockId: clause,
      find: "This",
      replace: "revised",
    },
  ]);
  apply("tracked-changes", [
    { id: "delete-closing", type: "deleteBlock", blockId: closing },
    { id: "merge-first", type: "mergeBlockWithNext", blockId: supplier, separator: " " },
  ]);
  const removedBreak = reviewer
    .getChanges()
    .find((change) => change.type === "paragraphMarkDeleted");
  if (!removedBreak) throw new Error("Missing pending paragraph break");
  expect(reviewer.rejectChange(removedBreak)).toBe(true);
  reviewer.acceptAll();
  apply("tracked-changes", [
    {
      id: "replace-clause",
      type: "replaceInBlock",
      blockId: clause,
      find: "below",
      replace: "the Customer",
    },
    { id: "merge-again", type: "mergeBlockWithNext", blockId: supplier, separator: " " },
  ]);
  const live = reviewer.getContent().find((block) => block.id === supplier)?.previewRuns;
  const saved = (await reopen(reviewer))
    .getContent()
    .find((block) => block.id === supplier)?.previewRuns;
  expect(live).toEqual(saved);
  expect(live?.at(-1)?.fontSizePt).toBe(11);
});
