/**
 * Resolving tracked work that builds on pending revisions of its own: the
 * sequences the property in `trackedResolution.property.test.ts` found, each
 * pinned as the smallest one that showed it.
 */

import { describe, expect, test } from "bun:test";

import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperationMode,
} from "../document-operations";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { FolioDocxReviewer } from "./headless";

const SUPPLIER = "The Supplier delivers the goods on time and in good order.";
const SIGNED = "Signed in two copies.";
const ORIGINAL = [
  "Service Agreement",
  "This agreement is made between the parties named below.",
  SUPPLIER,
  "The Buyer pays each invoice within thirty days.",
  SIGNED,
];

const open = async (bytes?: ArrayBuffer): Promise<FolioDocxReviewer> => {
  if (bytes) {
    return FolioDocxReviewer.fromBuffer(bytes, { author: "Reviewer" });
  }
  const document = fromMarkdown(["# Service Agreement", ...ORIGINAL.slice(1)].join("\n\n"));
  const { docx } = await ensureParaIds(new Uint8Array(await createDocx(document)));
  return FolioDocxReviewer.fromBuffer(
    docx.buffer.slice(docx.byteOffset, docx.byteOffset + docx.byteLength) as ArrayBuffer,
    { author: "Reviewer" },
  );
};

const reopen = async (reviewer: FolioDocxReviewer): Promise<FolioDocxReviewer> =>
  open(await reviewer.toBuffer());

const texts = (reviewer: FolioDocxReviewer): string[] =>
  reviewer.getContent().map(({ text }) => text);

const idOf = (reviewer: FolioDocxReviewer, prefix: string): string => {
  const block = reviewer.getContent().find(({ text }) => text.startsWith(prefix));
  if (!block) {
    throw new Error(`no block starts with "${prefix}"`);
  }
  return block.id;
};

const applier =
  (reviewer: FolioDocxReviewer, mode: FolioDocumentOperationMode = "tracked-changes") =>
  (operation: Record<string, unknown>): void => {
    const result = reviewer.applyDocumentOperations({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode,
      operations: [{ id: "1", ...operation }],
    } as never);
    expect(result.applied).toHaveLength(1);
  };

describe("a tracked merge of a break that is itself a pending insertion", () => {
  // Insert after the last paragraph, split that paragraph, and merge the
  // split's second half into the insertion: the merged break is the one the
  // insertion added.
  const build = async (): Promise<FolioDocxReviewer> => {
    const reviewer = await open();
    const apply = applier(reviewer);
    apply({
      type: "insertAfterBlock",
      blockId: idOf(reviewer, "Signed"),
      text: "Inserted clause.",
    });
    apply({ type: "splitBlock", blockId: idOf(reviewer, "Signed"), offset: "Signed in ".length });
    apply({ type: "mergeBlockWithNext", blockId: idOf(reviewer, "two copies"), separator: " " });
    return reviewer;
  };

  test("retracts that insertion: the paragraphs join at once", async () => {
    expect(texts(await build()).slice(-2)).toEqual(["Signed in ", "two copies. Inserted clause."]);
  });

  test("rejecting everything gives the original back", async () => {
    const reviewer = await build();
    reviewer.rejectAll();
    expect(texts(reviewer)).toEqual(ORIGINAL);
    expect(texts(await reopen(reviewer))).toEqual(ORIGINAL);
  });

  test("accepting everything keeps the split and the merge", async () => {
    const reviewer = await build();
    reviewer.acceptAll();
    expect(texts(reviewer).slice(-2)).toEqual(["Signed in ", "two copies. Inserted clause."]);
  });

  test("merging the last paragraph into an insertion after it rejects to the original", async () => {
    const reviewer = await open();
    const apply = applier(reviewer);
    apply({ type: "insertAfterBlock", blockId: idOf(reviewer, "Signed"), text: "Inserted." });
    apply({ type: "mergeBlockWithNext", blockId: idOf(reviewer, "Signed"), separator: " " });
    reviewer.rejectAll();
    expect(texts(reviewer)).toEqual(ORIGINAL);
  });

  test("merging a split's first half into its deleted second half rejects to the original", async () => {
    const reviewer = await open();
    const apply = applier(reviewer);
    apply({ type: "splitBlock", blockId: idOf(reviewer, "The Supplier"), offset: 13 });
    apply({ type: "deleteBlock", blockId: idOf(reviewer, "delivers") });
    apply({ type: "mergeBlockWithNext", blockId: idOf(reviewer, "The Supplier"), separator: " " });
    reviewer.rejectAll();
    expect(texts(reviewer)).toEqual(ORIGINAL);
  });
});

describe("rejecting a split with an inserted table after its first half", () => {
  test("joins the halves again once the table is gone", async () => {
    const reviewer = await open();
    const apply = applier(reviewer);
    apply({ type: "splitBlock", blockId: idOf(reviewer, "The Supplier"), offset: 47 });
    apply({
      type: "insertTable",
      blockId: idOf(reviewer, "The Supplier"),
      rows: [["Term", "Value"]],
    });
    expect(texts(reviewer)).toContain("Term");
    reviewer.rejectAll();
    expect(texts(reviewer)).toEqual(ORIGINAL);
  });
});
