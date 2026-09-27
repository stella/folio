/**
 * Edits that name what an earlier tracked edit left pending.
 *
 * A paragraph whose text and mark are pending deletion reads as a blank
 * block. Word keeps text typed into it joined to the next paragraph, so a
 * `replaceBlock` there would run its text into that paragraph once accepted;
 * it is refused (`pendingDeletion`) in every mode, and the document stays as
 * it was. The same holds for a paragraph in a table row pending deletion.
 *
 * A tracked deletion over pending inserted text saves as `w:ins > w:del`,
 * which reopens as one deletion mark recording the insertion around it; the
 * reader lists both revisions either way.
 */

import { describe, expect, test } from "bun:test";

import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { paragraph, table } from "../docx/server/build";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { fromMarkdown } from "../markdown";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";
import { hashFolioAIBlockText } from "./snapshot";
import type { FolioAIEditApplyMode } from "./types";

const SUPPLIER = "The Supplier delivers the goods on time and in good order.";
const BUYER = "The Buyer pays each invoice within thirty days.";
const TEXTS = [
  "Service Agreement",
  "This agreement is made between the parties named below.",
  SUPPLIER,
  BUYER,
  "Signed in two copies.",
];

const toArrayBuffer = (bytes: Uint8Array): ArrayBuffer =>
  bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;

const plainDocument = async (): Promise<ArrayBuffer> => {
  const document = fromMarkdown(`# ${TEXTS.join("\n\n")}`);
  return toArrayBuffer((await ensureParaIds(new Uint8Array(await createDocx(document)))).docx);
};

const tableDocument = async (): Promise<ArrayBuffer> => {
  const document = createEmptyDocument();
  document.package.document.content = [
    paragraph("Prices"),
    table({ header: ["Item", "Price"], rows: [["Goods", "100"]] }),
    paragraph("Taxes are extra."),
  ];
  return toArrayBuffer((await ensureParaIds(new Uint8Array(await createDocx(document)))).docx);
};

const open = (buffer: ArrayBuffer) => FolioDocxReviewer.fromBuffer(buffer, { author: "Reviewer" });
const reopen = async (reviewer: FolioDocxReviewer) => open(await reviewer.toBuffer());

const apply = (
  reviewer: FolioDocxReviewer,
  mode: FolioAIEditApplyMode,
  operations: FolioDocumentOperation[],
) =>
  reviewer.applyDocumentOperations({
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    mode,
    operations,
  });

const textsOf = (reviewer: FolioDocxReviewer) => reviewer.getContent().map(({ text }) => text);

const blockId = (reviewer: FolioDocxReviewer, text: string): string => {
  const block = reviewer.getContent().find((candidate) => candidate.text === text);
  if (!block) {
    throw new Error(`no block reads "${text}"`);
  }
  return block.id;
};

/** Every change accepted, or rejected, in a copy saved and reopened from `reviewer`. */
const resolved = async (reviewer: FolioDocxReviewer, resolution: "accept" | "reject") => {
  const copy = await reopen(reviewer);
  if (resolution === "accept") copy.acceptAll();
  else copy.rejectAll();
  return textsOf(await reopen(copy));
};

const MODES = ["direct", "tracked-changes", "suggested"] as const;

describe("replaceBlock on a paragraph pending deletion", () => {
  for (const saved of [true, false]) {
    for (const mode of MODES) {
      test(`is refused and changes nothing (${mode}, deletion ${saved ? "saved and reopened" : "made in this session"})`, async () => {
        let reviewer = await open(await plainDocument());
        const deleted = apply(reviewer, "tracked-changes", [
          { id: "delete", type: "deleteBlock", blockId: blockId(reviewer, SUPPLIER) },
        ]);
        expect(deleted.applied).toHaveLength(1);
        if (saved) {
          reviewer = await reopen(reviewer);
        }
        const blank = reviewer.getContent().find((block) => block.text === "");
        expect(blank).toBeDefined();
        const before = await reviewer.toBuffer();
        const changesBefore = reviewer.getChanges();

        const result = apply(reviewer, mode, [
          { id: "rewrite", type: "replaceBlock", blockId: blank?.id ?? "", text: "Rewritten." },
        ]);

        expect(result.applied).toEqual([]);
        expect(result.skipped).toEqual([{ id: "rewrite", reason: "pendingDeletion" }]);
        expect(result.issues).toEqual([
          expect.objectContaining({
            operationId: "rewrite",
            code: "pendingDeletion",
            retryable: false,
            recovery: "resolveTrackedChange",
          }),
        ]);
        expect(reviewer.getChanges()).toEqual(changesBefore);
        const unchanged = await open(before);
        expect(await resolved(reviewer, "accept")).toEqual(await resolved(unchanged, "accept"));
        expect(await resolved(reviewer, "accept")).toEqual(
          TEXTS.filter((text) => text !== SUPPLIER),
        );
        expect(await resolved(reviewer, "reject")).toEqual(TEXTS);
      });
    }
  }

  test("the heading deleted earlier in the session is refused too (collision seed 1088)", async () => {
    const reviewer = await open(await plainDocument());
    const heading = blockId(reviewer, "Service Agreement");
    apply(reviewer, "tracked-changes", [{ id: "delete", type: "deleteBlock", blockId: heading }]);
    const result = apply(reviewer, "tracked-changes", [
      { id: "rewrite", type: "replaceBlock", blockId: heading, text: "Written agreement." },
    ]);
    expect(result.skipped).toEqual([{ id: "rewrite", reason: "pendingDeletion" }]);
    expect(await resolved(reviewer, "accept")).toEqual(TEXTS.slice(1));
  });

  test("a paragraph whose mark alone is pending deletion is still rewritten", async () => {
    const reviewer = await open(await plainDocument());
    const supplier = blockId(reviewer, SUPPLIER);
    apply(reviewer, "tracked-changes", [
      { id: "merge", type: "mergeBlockWithNext", blockId: supplier, separator: " " },
    ]);
    const result = apply(reviewer, "tracked-changes", [
      { id: "rewrite", type: "replaceBlock", blockId: supplier, text: "Rewritten." },
    ]);
    expect(result.applied.map(({ id }) => id)).toEqual(["rewrite"]);
    // The block keeps its content, so it is not pending deletion: the merge
    // the document already asks for still joins it to the next paragraph.
    expect(await resolved(reviewer, "accept")).toEqual([
      ...TEXTS.slice(0, 2),
      `Rewritten.${BUYER}`,
      ...TEXTS.slice(4),
    ]);
    expect(await resolved(reviewer, "reject")).toEqual(TEXTS);
  });

  for (const mode of MODES) {
    test(`in a table row pending deletion is refused (${mode})`, async () => {
      let reviewer = await open(await tableDocument());
      const deleted = apply(reviewer, "tracked-changes", [
        { id: "delete", type: "deleteTableRow", blockId: blockId(reviewer, "Goods") },
      ]);
      expect(deleted.applied).toHaveLength(1);
      reviewer = await reopen(reviewer);
      const cell = reviewer
        .getContent()
        .find((block) => block.text === "" && block.table?.rowIndex === 1);
      expect(cell).toBeDefined();

      const result = apply(reviewer, mode, [
        { id: "rewrite", type: "replaceBlock", blockId: cell?.id ?? "", text: "Services" },
      ]);

      expect(result.skipped).toEqual([{ id: "rewrite", reason: "pendingDeletion" }]);
      expect(await resolved(reviewer, "accept")).toEqual([
        "Prices",
        "Item",
        "Price",
        "Taxes are extra.",
      ]);
    });
  }
});

describe("a tracked deletion over a saved pending insertion", () => {
  const changeKinds = (reviewer: FolioDocxReviewer) =>
    [...new Set(reviewer.getChanges().map(({ type, author }) => `${type} by ${author}`))].sort();

  for (const formatFirst of [true, false]) {
    test(`reads back the insertion it deletes (${formatFirst ? "after a format and a replace" : "after a replace"}; collision seed 1185)`, async () => {
      const first = await open(await plainDocument());
      const supplier = blockId(first, SUPPLIER);
      const edits: FolioDocumentOperation[] = [
        {
          id: "replace",
          type: "replaceInBlock",
          blockId: supplier,
          find: "The",
          replace: "changed",
        },
      ];
      if (formatFirst) {
        edits.unshift({
          id: "format",
          type: "formatRange",
          range: {
            type: "textRange",
            story: "main",
            blockId: supplier,
            startOffset: 0,
            endOffset: 3,
            selectedTextHash: hashFolioAIBlockText("The"),
          },
          formatting: { bold: true },
        });
      }
      expect(apply(first, "tracked-changes", edits).applied).toHaveLength(edits.length);

      const reviewer = await reopen(first);
      expect(
        apply(reviewer, "tracked-changes", [
          { id: "delete", type: "deleteBlock", blockId: supplier },
        ]).applied,
      ).toHaveLength(1);
      const live = reviewer.getChanges();
      expect(live).toContainEqual(expect.objectContaining({ type: "insertion", text: "changed" }));

      const reopened = await reopen(reviewer);
      expect(changeKinds(reopened)).toEqual(changeKinds(reviewer));
      expect(reopened.getChanges()).toContainEqual(
        expect.objectContaining({ type: "insertion", text: "changed" }),
      );
      expect(await resolved(reopened, "accept")).toEqual(await resolved(reviewer, "accept"));
      expect(await resolved(reopened, "reject")).toEqual(TEXTS);

      // The listed insertion resolves on its own: rejecting it takes the
      // inserted word away, deletion and all.
      const insertion = reopened.getChanges().find((change) => change.type === "insertion");
      expect(insertion && reopened.rejectChange(insertion)).toBe(true);
      expect(reopened.getChanges().some((change) => change.type === "insertion")).toBe(false);
      expect(reopened.getContent().find((block) => block.id === supplier)?.text).toBe("");
      reopened.rejectAll();
      expect(textsOf(reopened)).toEqual(TEXTS);
    });
  }
});
