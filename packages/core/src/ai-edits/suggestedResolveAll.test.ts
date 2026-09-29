/**
 * `"suggested"` edits stay out of a saved package until explicitly accepted.
 * Bulk acceptance resolves ordinary revisions while leaving them staged.
 */

import { describe, expect, test } from "bun:test";

import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { paragraph } from "../docx/server/build";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";

const buildDocument = async (): Promise<ArrayBuffer> => {
  const document = createEmptyDocument();
  document.package.document.content = [
    paragraph("First clause."),
    paragraph("Signed in two copies."),
  ];
  const { docx } = await ensureParaIds(new Uint8Array(await createDocx(document)));
  return docx.buffer.slice(docx.byteOffset, docx.byteOffset + docx.byteLength) as ArrayBuffer;
};

const suggest = async (): Promise<FolioDocxReviewer> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(await buildDocument(), { author: "AI" });
  const anchor = reviewer.getContent().find(({ text }) => text === "Signed in two copies.");
  if (!anchor) {
    throw new Error("the fixture paragraph is missing");
  }
  const result = reviewer.applyDocumentOperations({
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    mode: "suggested",
    operations: [
      { id: "insert", type: "insertAfterBlock", blockId: anchor.id, text: "Suggested clause." },
      { id: "replace", type: "replaceInBlock", blockId: anchor.id, find: "two", replace: "three" },
    ],
  });
  expect(result.applied).toHaveLength(2);
  return reviewer;
};

const savedTexts = async (reviewer: FolioDocxReviewer): Promise<string[]> =>
  (await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer()))
    .getContent()
    .map(({ text }) => text);

describe("resolving every suggestion headlessly", () => {
  test("a save before resolving keeps the document as it was", async () => {
    expect(await savedTexts(await suggest())).toEqual(["First clause.", "Signed in two copies."]);
  });

  test("a suggested block deletion leaves no tracked change in the saved package", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await buildDocument(), { author: "AI" });
    const target = reviewer.getContent().find(({ text }) => text === "First clause.");
    if (!target) {
      throw new Error("the fixture paragraph is missing");
    }
    const result = reviewer.applyDocumentOperations({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "suggested",
      operations: [{ id: "delete", type: "deleteBlock", blockId: target.id }],
    });
    expect(result.applied).toHaveLength(1);
    const saved = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
    expect(saved.getChanges()).toEqual([]);
    expect(saved.getContent().map(({ text }) => text)).toEqual([
      "First clause.",
      "Signed in two copies.",
    ]);
    expect(reviewer.acceptAll()).toBe(0);
    expect(reviewer.acceptSuggestion("delete")).toBe(true);
    expect(reviewer.acceptAll()).toBeGreaterThan(0);
    expect(await savedTexts(reviewer)).toEqual(["Signed in two copies."]);
  });

  test("a suggested deletion cannot retract an ordinary tracked split", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await buildDocument(), { author: "AI" });
    const target = reviewer.getContent().find(({ text }) => text === "Signed in two copies.");
    if (!target) throw new Error("fixture paragraph missing");
    const split = reviewer.applyDocumentOperations({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      operations: [{ id: "split", type: "splitBlock", blockId: target.id, offset: 10 }],
    });
    expect(split.applied).toHaveLength(1);
    const firstHalf = reviewer.getContent().find(({ text }) => text === "Signed in ");
    if (!firstHalf) throw new Error("split paragraph missing");
    const deletion = reviewer.applyDocumentOperations({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "suggested",
      operations: [{ id: "delete", type: "deleteBlock", blockId: firstHalf.id }],
    });
    expect(deletion.applied).toEqual([]);
    expect(deletion.skipped).toEqual([{ id: "delete", reason: "unsupportedMode" }]);
    expect(await savedTexts(reviewer)).toEqual(["First clause.", "Signed in ", "two copies."]);
  });

  test("acceptAll leaves suggested edits staged and out of the saved package", async () => {
    const reviewer = await suggest();
    const before = reviewer.getChanges();
    expect(reviewer.acceptAll()).toBe(0);
    expect(reviewer.getChanges()).toEqual(before);
    expect(reviewer.getContent().map(({ text }) => text)).toEqual([
      "First clause.",
      "Signed in three copies.",
      "Suggested clause.",
    ]);
    expect(await savedTexts(reviewer)).toEqual(["First clause.", "Signed in two copies."]);
  });

  test("explicit acceptance after bulk acceptance persists the suggestion", async () => {
    const reviewer = await suggest();
    expect(reviewer.acceptAll()).toBe(0);
    expect(reviewer.acceptSuggestion("insert")).toBe(true);
    expect(reviewer.acceptSuggestion("replace")).toBe(true);
    expect(await savedTexts(reviewer)).toEqual([
      "First clause.",
      "Signed in three copies.",
      "Suggested clause.",
    ]);
  });

  test("acceptAll resolves ordinary revisions while preserving a pending suggestion", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await buildDocument(), { author: "AI" });
    const first = reviewer.getContent().find(({ text }) => text === "First clause.");
    const last = reviewer.getContent().find(({ text }) => text === "Signed in two copies.");
    if (!first || !last) throw new Error("the fixture paragraphs are missing");
    const tracked = reviewer.applyDocumentOperations({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      operations: [
        {
          id: "tracked",
          type: "replaceInBlock",
          blockId: first.id,
          find: "First",
          replace: "Opening",
        },
      ],
    });
    expect(tracked.applied).toHaveLength(1);
    const suggested = reviewer.applyDocumentOperations({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "suggested",
      operations: [
        { id: "pending", type: "replaceInBlock", blockId: last.id, find: "two", replace: "three" },
      ],
    });
    expect(suggested.applied).toHaveLength(1);

    expect(reviewer.acceptAll()).toBe(2);
    expect(reviewer.getChanges()).toHaveLength(2);
    expect(await savedTexts(reviewer)).toEqual(["Opening clause.", "Signed in two copies."]);
    const records = JSON.parse(JSON.stringify(reviewer.exportPendingSuggestions()));
    const reopened = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer(), {
      author: "AI",
    });
    expect(reopened.loadPendingSuggestions(records)).toEqual([
      { status: "restaged", suggestionId: "pending" },
    ]);
    expect(
      reopened.getContent().map(({ kind, text, previewRuns }) => ({ kind, text, previewRuns })),
    ).toEqual(
      reviewer.getContent().map(({ kind, text, previewRuns }) => ({ kind, text, previewRuns })),
    );
    expect(reviewer.acceptSuggestion("pending")).toBe(true);
    expect(await savedTexts(reviewer)).toEqual(["Opening clause.", "Signed in three copies."]);
  });

  test("a suggested deletion cannot overwrite a tracked paragraph revision", async () => {
    const document = fromMarkdown("# Delivery Terms\n\nThe following apply to every order.");
    const { docx } = await ensureParaIds(new Uint8Array(await createDocx(document)));
    const reviewer = await FolioDocxReviewer.fromBuffer(docx, { author: "AI" });
    const [first, last] = reviewer.getContent();
    if (!first || !last) throw new Error("fixture paragraphs missing");
    expect(
      reviewer.applyDocumentOperations({
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "tracked-changes",
        operations: [
          {
            id: "tracked",
            type: "replaceInBlock",
            blockId: last.id,
            find: "every",
            replace: "each",
          },
        ],
      }).applied,
    ).toHaveLength(1);
    const result = reviewer.applyDocumentOperations({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "suggested",
      operations: [
        {
          id: "pending",
          type: "replaceInBlock",
          blockId: first.id,
          find: "Delivery",
          replace: "revised",
        },
        { id: "delete", type: "deleteBlock", blockId: last.id },
      ],
    });
    expect(result.applied.map(({ id }) => id)).toEqual(["pending"]);
    expect(result.skipped).toEqual([{ id: "delete", reason: "unsupportedMode" }]);
    reviewer.acceptAll();
    expect(await savedTexts(reviewer)).toEqual([
      "Delivery Terms",
      "The following apply to each order.",
    ]);
  });

  test("acceptAll fails without changing the document when a suggestion cannot be restaged", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await buildDocument(), { author: "AI" });
    const [first, last] = reviewer.getContent();
    if (!first || !last) throw new Error("fixture paragraphs missing");
    expect(
      reviewer.applyDocumentOperations({
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "tracked-changes",
        operations: [
          {
            id: "tracked",
            type: "replaceInBlock",
            blockId: first.id,
            find: "First",
            replace: "Opening",
          },
        ],
      }).applied,
    ).toHaveLength(1);
    expect(
      reviewer.applyDocumentOperations({
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "suggested",
        operations: [
          {
            id: "pending",
            type: "replaceInBlock",
            blockId: last.id,
            find: "two",
            replace: "three",
          },
        ],
      }).applied,
    ).toHaveLength(1);
    const registry: unknown = Reflect.get(reviewer, "pendingSuggestions");
    if (
      typeof registry !== "object" ||
      registry === null ||
      !("clear" in registry) ||
      typeof registry.clear !== "function"
    ) {
      throw new Error("the pending registry is missing");
    }
    registry.clear();
    const before = reviewer.getContent();
    const changes = reviewer.getChanges();
    expect(() => reviewer.acceptAll()).toThrow(
      "A pending suggestion has no operation to restage after bulk acceptance",
    );
    expect(reviewer.getContent()).toEqual(before);
    expect(reviewer.getChanges()).toEqual(changes);
  });

  test("acceptAll refreshes multiple pending records after a separate tracked edit", async () => {
    const reviewer = await suggest();
    const first = reviewer.getContent()[0];
    if (!first) throw new Error("fixture paragraph missing");
    expect(
      reviewer.applyDocumentOperations({
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "tracked-changes",
        operations: [
          {
            id: "tracked",
            type: "replaceInBlock",
            blockId: first.id,
            find: "First",
            replace: "Opening",
          },
        ],
      }).applied,
    ).toHaveLength(1);
    expect(reviewer.acceptAll()).toBe(2);
    const records = JSON.parse(JSON.stringify(reviewer.exportPendingSuggestions()));
    expect(records).toHaveLength(2);
    const reopened = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer(), {
      author: "AI",
    });
    expect(reopened.loadPendingSuggestions(records)).toEqual([
      { status: "restaged", suggestionId: "insert" },
      { status: "restaged", suggestionId: "replace" },
    ]);
    expect(
      reopened.getContent().map(({ kind, text, previewRuns }) => ({ kind, text, previewRuns })),
    ).toEqual(
      reviewer.getContent().map(({ kind, text, previewRuns }) => ({ kind, text, previewRuns })),
    );
  });

  test("bulk acceptance preserves a pending comment and its comment id", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await buildDocument());
    const [first, last] = reviewer.getContent();
    if (!first || !last) throw new Error("fixture paragraphs missing");
    expect(
      reviewer.applyDocumentOperations({
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "tracked-changes",
        operations: [
          {
            id: "tracked",
            type: "replaceInBlock",
            blockId: first.id,
            find: "First",
            replace: "Opening",
          },
        ],
      }).applied,
    ).toHaveLength(1);
    const comment = reviewer.applyDocumentOperations({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "suggested",
      operations: [
        {
          id: "pending-comment",
          type: "replaceInBlock",
          blockId: last.id,
          find: "two",
          replace: "three",
          comment: { text: "Review this." },
        },
      ],
    });
    expect(comment.skipped).toEqual([]);
    expect(comment.applied).toHaveLength(1);
    const commentId = comment.applied.at(0)?.commentId;
    expect(commentId).toBeDefined();
    expect(reviewer.acceptAll()).toBeGreaterThan(0);
    expect(reviewer.exportPendingSuggestions().at(0)?.commentId).toBe(commentId);
    const records = JSON.parse(JSON.stringify(reviewer.exportPendingSuggestions()));
    const reopened = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
    expect(reopened.getComments()).toEqual([]);
    expect(reopened.loadPendingSuggestions(records)).toEqual([
      { status: "restaged", suggestionId: "pending-comment" },
    ]);
    expect(reopened.acceptAll()).toBe(0);
  });

  test("bulk acceptance drops a proposal whose anchor was merged away", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await buildDocument());
    const [first, last] = reviewer.getContent();
    if (!first || !last) throw new Error("fixture paragraphs missing");
    // Accepting the merge removes the first paragraph's mark, and the
    // paragraph left is the one whose mark stays: the first is gone.
    expect(
      reviewer.applyDocumentOperations({
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "suggested",
        operations: [
          {
            id: "pending",
            type: "replaceInBlock",
            blockId: first.id,
            find: "First",
            replace: "Opening",
          },
        ],
      }).applied,
    ).toHaveLength(1);
    expect(
      reviewer.applyDocumentOperations({
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "tracked-changes",
        operations: [{ id: "merge", type: "mergeBlockWithNext", blockId: first.id }],
      }).applied,
    ).toHaveLength(1);
    expect(reviewer.acceptAll()).toBeGreaterThan(0);
    expect(reviewer.exportPendingSuggestions()).toEqual([]);
  });

  test("refuses a suggestion that would retract a tracked insertion", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await buildDocument(), { author: "AI" });
    const first = reviewer.getContent()[0];
    if (!first) throw new Error("the fixture paragraph is missing");
    expect(
      reviewer.applyDocumentOperations({
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "tracked-changes",
        operations: [
          { id: "tracked", type: "insertAfterBlock", blockId: first.id, text: "Added." },
        ],
      }).applied,
    ).toHaveLength(1);
    const added = reviewer.getContent().find(({ text }) => text === "Added.");
    if (!added) throw new Error("the tracked insertion is missing");
    const result = reviewer.applyDocumentOperations({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "suggested",
      operations: [{ id: "pending", type: "deleteBlock", blockId: added.id }],
    });
    expect(result.applied).toEqual([]);
    expect(result.skipped).toEqual([{ id: "pending", reason: "unsupportedMode" }]);
    reviewer.acceptAll();
    expect(await savedTexts(reviewer)).toEqual([
      "First clause.",
      "Added.",
      "Signed in two copies.",
    ]);
  });

  test("refuses suggestions anchored to a tracked deletion", async () => {
    const writer = await FolioDocxReviewer.fromBuffer(await buildDocument(), { author: "AI" });
    const first = writer.getContent()[0];
    if (!first) throw new Error("fixture paragraph missing");
    expect(
      writer.applyDocumentOperations({
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "tracked-changes",
        operations: [{ id: "tracked", type: "deleteBlock", blockId: first.id }],
      }).applied,
    ).toHaveLength(1);
    const bytes = await writer.toBuffer();
    for (const type of ["insertAfterBlock", "insertBeforeBlock"] as const) {
      const baseline = await FolioDocxReviewer.fromBuffer(bytes, { author: "AI" });
      const staged = await FolioDocxReviewer.fromBuffer(bytes, { author: "AI" });
      const blank = staged.getContent().find(({ text }) => text === "");
      if (!blank) throw new Error("deleted paragraph missing");
      const result = staged.applyDocumentOperations({
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "suggested",
        operations: [{ id: "pending", type, blockId: blank.id, text: "Added." }],
      });
      expect(result.applied).toEqual([]);
      expect(result.skipped).toEqual([{ id: "pending", reason: "pendingDeletion" }]);
      const beforeBulk = await FolioDocxReviewer.fromBuffer(await staged.toBuffer());
      expect(beforeBulk.getContent()).toEqual(baseline.getContent());
      baseline.acceptAll();
      staged.acceptAll();
      expect((await FolioDocxReviewer.fromBuffer(await staged.toBuffer())).getContent()).toEqual(
        (await FolioDocxReviewer.fromBuffer(await baseline.toBuffer())).getContent(),
      );
    }
  });

  test("suggested replacement and deletion preserve the saved heading", async () => {
    const doc = fromMarkdown(
      "# Delivery Terms\n\nThe following apply to every order.\n\n- Goods are packed securely\n- Goods are insured in transit\n\nPayment happens in stages.\n\n1. Deposit on signature\n2. Balance on delivery\n3. Retention after inspection\n\nClosing remarks.",
    );
    const { docx } = await ensureParaIds(new Uint8Array(await createDocx(doc)));
    const reviewer = await FolioDocxReviewer.fromBuffer(
      docx.buffer.slice(docx.byteOffset, docx.byteOffset + docx.byteLength) as ArrayBuffer,
      { author: "AI" },
    );
    const block = (prefix: string) =>
      reviewer.getContent().find(({ text }) => text.startsWith(prefix));
    const opening = block("The following");
    const payment = block("Payment happens");
    const insured = block("Goods are insured");
    if (!opening || !payment || !insured) throw new Error("fixture block missing");
    const tracked = reviewer.applyDocumentOperations({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      operations: [
        { id: "t1", type: "replaceInBlock", blockId: opening.id, find: "every", replace: "each" },
        {
          id: "t2",
          type: "insertAfterBlock",
          blockId: payment.id,
          text: "Stages are invoiced separately.",
        },
        { id: "t3", type: "deleteBlock", blockId: insured.id },
      ],
    });
    expect(tracked.applied).toHaveLength(3);
    const trackedReopened = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer(), {
      author: "AI",
    });
    const baseline = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer(), {
      author: "AI",
    });
    baseline.acceptAll();
    const baselineContent = (
      await FolioDocxReviewer.fromBuffer(await baseline.toBuffer())
    ).getContent();
    const heading = trackedReopened.getContent()[0];
    if (!heading) throw new Error("heading missing");
    const result = trackedReopened.applyDocumentOperations({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "suggested",
      operations: [
        {
          id: "replace",
          type: "replaceInBlock",
          blockId: heading.id,
          find: "Delivery",
          replace: "revised",
        },
        { id: "delete", type: "deleteBlock", blockId: heading.id },
      ],
    });
    expect(result.applied.map(({ id }) => id)).toEqual(["replace"]);
    expect(result.skipped.map(({ id }) => id)).toEqual(["delete"]);
    trackedReopened.acceptAll();
    const reopened = await FolioDocxReviewer.fromBuffer(await trackedReopened.toBuffer());
    expect(reopened.getContent()).toEqual(baselineContent);
  });

  test("acceptChange on a suggested paragraph insert keeps the paragraph through the save", async () => {
    const reviewer = await suggest();
    const insertion = reviewer
      .getChanges()
      .find(({ type, text }) => type === "insertion" && text === "Suggested clause.");
    if (!insertion) {
      throw new Error("the suggested insertion is not listed");
    }
    expect(reviewer.acceptChange(insertion)).toBe(true);
    expect(await savedTexts(reviewer)).toContain("Suggested clause.");
  });

  test("rejectChange on a suggested paragraph insert removes the paragraph", async () => {
    const reviewer = await suggest();
    const insertion = reviewer
      .getChanges()
      .find(({ type, text }) => type === "insertion" && text === "Suggested clause.");
    if (!insertion) {
      throw new Error("the suggested insertion is not listed");
    }
    expect(reviewer.rejectChange(insertion)).toBe(true);
    expect(reviewer.getContent().map(({ text }) => text)).not.toContain("");
    expect(reviewer.getContent().map(({ text }) => text)).not.toContain("Suggested clause.");
  });

  test("rejectAll removes every suggestion", async () => {
    const reviewer = await suggest();
    reviewer.rejectAll();
    expect(reviewer.getContent().map(({ text }) => text)).toEqual([
      "First clause.",
      "Signed in two copies.",
    ]);
    expect(await savedTexts(reviewer)).toEqual(["First clause.", "Signed in two copies."]);
  });
});
