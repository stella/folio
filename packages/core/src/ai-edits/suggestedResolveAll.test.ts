/**
 * `"suggested"` edits stay out of a saved package until explicitly accepted.
 * Bulk acceptance resolves ordinary revisions while leaving them staged.
 */

import { describe, expect, test } from "bun:test";

import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { paragraph } from "../docx/server/build";
import { fromMarkdown } from "../markdown/fromMarkdown";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";
import { hashFolioAIBlockText } from "./snapshot";

const buildDocument = async (): Promise<ArrayBuffer> => {
  const document = createEmptyDocument();
  document.package.document.content = [
    paragraph("First clause."),
    paragraph("Signed in two copies."),
  ];
  const { docx } = await ensureParaIds(
    new Uint8Array(await createDocx(document)),
  );
  return docx.buffer.slice(
    docx.byteOffset,
    docx.byteOffset + docx.byteLength,
  ) as ArrayBuffer;
};

const suggest = async (): Promise<FolioDocxReviewer> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(await buildDocument(), {
    author: "AI",
  });
  const anchor = reviewer
    .getContent()
    .find(({ text }) => text === "Signed in two copies.");
  if (!anchor) {
    throw new Error("the fixture paragraph is missing");
  }
  const result = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "suggested",
      operations: [
        {
          id: "insert",
          type: "insertAfterBlock",
          blockId: anchor.id,
          text: "Suggested clause.",
        },
        {
          id: "replace",
          type: "replaceInBlock",
          blockId: anchor.id,
          find: "two",
          replace: "three",
        },
      ],
    },
    { undefinedReferences: "refuse" },
  );
  expect(result.applied).toHaveLength(2);
  return reviewer;
};

const savedTexts = async (reviewer: FolioDocxReviewer): Promise<string[]> =>
  (await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer()))
    .getContent()
    .map(({ text }) => text);

const SESSIONS = ["legacy", "canonical"] as const;

const documentOf = async (texts: readonly string[]): Promise<ArrayBuffer> => {
  const document = createEmptyDocument();
  document.package.document.content = texts.map((text) => paragraph(text));
  const { docx } = await ensureParaIds(
    new Uint8Array(await createDocx(document)),
  );
  return docx.buffer.slice(
    docx.byteOffset,
    docx.byteOffset + docx.byteLength,
  ) as ArrayBuffer;
};

const applyOne = (
  reviewer: FolioDocxReviewer,
  mode: "suggested" | "tracked-changes",
  operation: FolioDocumentOperation,
): void => {
  expect(
    reviewer.applyDocumentOperations(
      {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode,
        operations: [operation],
      },
      { undefinedReferences: "refuse" },
    ).applied,
  ).toHaveLength(1);
};

/**
 * A tracked edit and a suggestion, with the reviewer's private registry
 * cleared: the document holds a suggestion no record describes.
 */
const withUnrecordedSuggestion = async (
  session: HeadlessReviewerSession,
): Promise<FolioDocxReviewer> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(await buildDocument(), {
    author: "AI",
    session,
  });
  const [first, last] = reviewer.getContent();
  if (!first || !last) throw new Error("fixture paragraphs missing");
  applyOne(reviewer, "tracked-changes", {
    id: "tracked",
    type: "replaceInBlock",
    blockId: first.id,
    find: "First",
    replace: "Opening",
  });
  applyOne(reviewer, "suggested", {
    id: "pending",
    type: "replaceInBlock",
    blockId: last.id,
    find: "two",
    replace: "three",
  });
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
  return reviewer;
};

describe("resolving every suggestion headlessly", () => {
  test("a save before resolving keeps the document as it was", async () => {
    expect(await savedTexts(await suggest())).toEqual([
      "First clause.",
      "Signed in two copies.",
    ]);
  });

  test("a suggested block deletion leaves no tracked change in the saved package", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await buildDocument(), {
      author: "AI",
    });
    const target = reviewer
      .getContent()
      .find(({ text }) => text === "First clause.");
    if (!target) {
      throw new Error("the fixture paragraph is missing");
    }
    const result = reviewer.applyDocumentOperations(
      {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "suggested",
        operations: [{ id: "delete", type: "deleteBlock", blockId: target.id }],
      },
      { undefinedReferences: "refuse" },
    );
    expect(result.applied).toHaveLength(1);
    const saved = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
    expect(saved.getChanges()).toEqual([]);
    expect(saved.getContent().map(({ text }) => text)).toEqual([
      "First clause.",
      "Signed in two copies.",
    ]);
    expect(reviewer.acceptAll().resolved).toBe(0);
    expect(reviewer.acceptSuggestion("delete")).toBe(true);
    expect(reviewer.acceptAll().resolved).toBeGreaterThan(0);
    expect(await savedTexts(reviewer)).toEqual(["Signed in two copies."]);
  });

  // The legacy engine refuses a suggested deletion over an ordinary tracked
  // record; the canonical session records the deletion as the suggestion's
  // own, beside the record it leaves alone (a pending record is taken back
  // only by its own author and suggestion).
  test.each(SESSIONS)(
    "a suggested deletion leaves an ordinary tracked split pending (%s session)",
    async (session) => {
      const reviewer = await FolioDocxReviewer.fromBuffer(
        await buildDocument(),
        { author: "AI", session },
      );
      const target = reviewer
        .getContent()
        .find(({ text }) => text === "Signed in two copies.");
      if (!target) throw new Error("fixture paragraph missing");
      const split = reviewer.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode: "tracked-changes",
          operations: [
            { id: "split", type: "splitBlock", blockId: target.id, offset: 10 },
          ],
        },
        { undefinedReferences: "refuse" },
      );
      expect(split.applied).toHaveLength(1);
      const firstHalf = reviewer
        .getContent()
        .find(({ text }) => text === "Signed in ");
      if (!firstHalf) throw new Error("split paragraph missing");
      const deletion = reviewer.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode: "suggested",
          operations: [
            { id: "delete", type: "deleteBlock", blockId: firstHalf.id },
          ],
        },
        { undefinedReferences: "refuse" },
      );
      if (session === "legacy") {
        expect(deletion.applied).toEqual([]);
        expect(deletion.skipped).toEqual([
          { id: "delete", reason: "unsupportedMode" },
        ]);
      } else {
        expect(deletion.applied.map(({ id }) => id)).toEqual(["delete"]);
        expect(deletion.skipped).toEqual([]);
      }
      expect(await savedTexts(reviewer)).toEqual([
        "First clause.",
        "Signed in ",
        "two copies.",
      ]);
    },
  );

  test("acceptAll leaves suggested edits staged and out of the saved package", async () => {
    const reviewer = await suggest();
    const before = reviewer.getChanges();
    expect(reviewer.acceptAll().resolved).toBe(0);
    expect(reviewer.getChanges()).toEqual(before);
    expect(reviewer.getContent().map(({ text }) => text)).toEqual([
      "First clause.",
      "Signed in three copies.",
      "Suggested clause.",
    ]);
    expect(await savedTexts(reviewer)).toEqual([
      "First clause.",
      "Signed in two copies.",
    ]);
  });

  test("explicit acceptance after bulk acceptance persists the suggestion", async () => {
    const reviewer = await suggest();
    expect(reviewer.acceptAll().resolved).toBe(0);
    expect(reviewer.acceptSuggestion("insert")).toBe(true);
    expect(reviewer.acceptSuggestion("replace")).toBe(true);
    expect(await savedTexts(reviewer)).toEqual([
      "First clause.",
      "Signed in three copies.",
      "Suggested clause.",
    ]);
  });

  test("acceptAll resolves ordinary revisions while preserving a pending suggestion", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await buildDocument(), {
      author: "AI",
    });
    const first = reviewer
      .getContent()
      .find(({ text }) => text === "First clause.");
    const last = reviewer
      .getContent()
      .find(({ text }) => text === "Signed in two copies.");
    if (!first || !last) throw new Error("the fixture paragraphs are missing");
    const tracked = reviewer.applyDocumentOperations(
      {
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
      },
      { undefinedReferences: "refuse" },
    );
    expect(tracked.applied).toHaveLength(1);
    const suggested = reviewer.applyDocumentOperations(
      {
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
      },
      { undefinedReferences: "refuse" },
    );
    expect(suggested.applied).toHaveLength(1);

    expect(reviewer.acceptAll().resolved).toBe(2);
    expect(reviewer.getChanges()).toHaveLength(2);
    expect(await savedTexts(reviewer)).toEqual([
      "Opening clause.",
      "Signed in two copies.",
    ]);
    const records = JSON.parse(
      JSON.stringify(reviewer.exportPendingSuggestions()),
    );
    const reopened = await FolioDocxReviewer.fromBuffer(
      await reviewer.toBuffer(),
      {
        author: "AI",
      },
    );
    expect(reopened.loadPendingSuggestions(records)).toEqual([
      { status: "restaged", suggestionId: "pending" },
    ]);
    expect(
      reopened
        .getContent()
        .map(({ kind, text, previewRuns }) => ({ kind, text, previewRuns })),
    ).toEqual(
      reviewer
        .getContent()
        .map(({ kind, text, previewRuns }) => ({ kind, text, previewRuns })),
    );
    expect(reviewer.acceptSuggestion("pending")).toBe(true);
    expect(await savedTexts(reviewer)).toEqual([
      "Opening clause.",
      "Signed in three copies.",
    ]);
  });

  // As above: the canonical session records the suggested deletion beside the
  // tracked replacement, which stays as it was.
  test.each(SESSIONS)(
    "a suggested deletion leaves a tracked paragraph revision pending (%s session)",
    async (session) => {
      const document = fromMarkdown(
        "# Delivery Terms\n\nThe following apply to every order.",
      );
      const { docx } = await ensureParaIds(
        new Uint8Array(await createDocx(document)),
      );
      const reviewer = await FolioDocxReviewer.fromBuffer(docx, {
        author: "AI",
        session,
      });
      const [first, last] = reviewer.getContent();
      if (!first || !last) throw new Error("fixture paragraphs missing");
      expect(
        reviewer.applyDocumentOperations(
          {
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
          },
          { undefinedReferences: "refuse" },
        ).applied,
      ).toHaveLength(1);
      const result = reviewer.applyDocumentOperations(
        {
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
        },
        { undefinedReferences: "refuse" },
      );
      if (session === "legacy") {
        expect(result.applied.map(({ id }) => id)).toEqual(["pending"]);
        expect(result.skipped).toEqual([
          { id: "delete", reason: "unsupportedMode" },
        ]);
      } else {
        expect(result.applied.map(({ id }) => id)).toEqual([
          "pending",
          "delete",
        ]);
        expect(result.skipped).toEqual([]);
      }
      reviewer.acceptAll();
      expect(await savedTexts(reviewer)).toEqual([
        "Delivery Terms",
        "The following apply to each order.",
      ]);
    },
  );

  // Legacy bulk acceptance rejects every proposal, resolves, then restages
  // each from its record; the session never restages.
  test("legacy acceptAll fails without changing the document when a suggestion cannot be restaged", async () => {
    const reviewer = await withUnrecordedSuggestion("legacy");
    const before = reviewer.getContent();
    const changes = reviewer.getChanges();
    expect(() => reviewer.acceptAll()).toThrow(
      "A pending suggestion has no operation to restage after bulk acceptance",
    );
    expect(reviewer.getContent()).toEqual(before);
    expect(reviewer.getChanges()).toEqual(changes);
  });

  for (const resolution of ["acceptAll", "rejectAll"] as const)
    test(`canonical ${resolution} panics on a suggestion the registry does not record`, async () => {
      const reviewer = await withUnrecordedSuggestion("canonical");
      const before = reviewer.getContent();
      const changes = reviewer.getChanges();
      expect(() => reviewer[resolution]()).toThrow(
        "A pending suggestion in the document has no record in the reviewer's registry",
      );
      expect(reviewer.getContent()).toEqual(before);
      expect(reviewer.getChanges()).toEqual(changes);
    });

  test("acceptAll refreshes multiple pending records after a separate tracked edit", async () => {
    const reviewer = await suggest();
    const first = reviewer.getContent()[0];
    if (!first) throw new Error("fixture paragraph missing");
    expect(
      reviewer.applyDocumentOperations(
        {
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
        },
        { undefinedReferences: "refuse" },
      ).applied,
    ).toHaveLength(1);
    expect(reviewer.acceptAll().resolved).toBe(2);
    const records = JSON.parse(
      JSON.stringify(reviewer.exportPendingSuggestions()),
    );
    expect(records).toHaveLength(2);
    const reopened = await FolioDocxReviewer.fromBuffer(
      await reviewer.toBuffer(),
      {
        author: "AI",
      },
    );
    expect(reopened.loadPendingSuggestions(records)).toEqual([
      { status: "restaged", suggestionId: "insert" },
      { status: "restaged", suggestionId: "replace" },
    ]);
    expect(
      reopened
        .getContent()
        .map(({ kind, text, previewRuns }) => ({ kind, text, previewRuns })),
    ).toEqual(
      reviewer
        .getContent()
        .map(({ kind, text, previewRuns }) => ({ kind, text, previewRuns })),
    );
  });

  test("bulk acceptance preserves a pending comment and its comment id", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await buildDocument());
    const [first, last] = reviewer.getContent();
    if (!first || !last) throw new Error("fixture paragraphs missing");
    expect(
      reviewer.applyDocumentOperations(
        {
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
        },
        { undefinedReferences: "refuse" },
      ).applied,
    ).toHaveLength(1);
    const comment = reviewer.applyDocumentOperations(
      {
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
      },
      { undefinedReferences: "refuse" },
    );
    expect(comment.skipped).toEqual([]);
    expect(comment.applied).toHaveLength(1);
    const commentId = comment.applied.at(0)?.commentId;
    expect(commentId).toBeDefined();
    expect(reviewer.acceptAll().resolved).toBeGreaterThan(0);
    expect(reviewer.exportPendingSuggestions().at(0)?.commentId).toBe(
      commentId,
    );
    const records = JSON.parse(
      JSON.stringify(reviewer.exportPendingSuggestions()),
    );
    const reopened = await FolioDocxReviewer.fromBuffer(
      await reviewer.toBuffer(),
    );
    expect(reopened.getComments()).toEqual([]);
    expect(reopened.loadPendingSuggestions(records)).toEqual([
      { status: "restaged", suggestionId: "pending-comment" },
    ]);
    expect(reopened.acceptAll().resolved).toBe(0);
  });

  // Legacy bulk acceptance restages proposals on the paragraphs they were made
  // on, so one whose paragraph was absorbed is lost (a legacy gap).
  test("legacy bulk acceptance drops a proposal whose anchor was merged away", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await buildDocument(), {
      session: "legacy",
    });
    const [first] = reviewer.getContent();
    if (!first) throw new Error("fixture paragraphs missing");
    applyOne(reviewer, "suggested", {
      id: "pending",
      type: "replaceInBlock",
      blockId: first.id,
      find: "First",
      replace: "Opening",
    });
    applyOne(reviewer, "tracked-changes", {
      id: "merge",
      type: "mergeBlockWithNext",
      blockId: first.id,
    });
    expect(reviewer.acceptAll().resolved).toBeGreaterThan(0);
    expect(reviewer.exportPendingSuggestions()).toEqual([]);
  });

  // Accepting the merge removes the first paragraph's mark; the paragraph
  // whose mark stays survives and holds the proposal, so its record names
  // the survivor and restages on the saved package.
  test("canonical bulk acceptance re-anchors a proposal whose anchor was merged away", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await buildDocument(), {
      session: "canonical",
    });
    const [first, last] = reviewer.getContent();
    if (!first || !last) throw new Error("fixture paragraphs missing");
    applyOne(reviewer, "suggested", {
      id: "pending",
      type: "replaceInBlock",
      blockId: first.id,
      find: "First",
      replace: "Opening",
    });
    applyOne(reviewer, "tracked-changes", {
      id: "merge",
      type: "mergeBlockWithNext",
      blockId: first.id,
    });
    expect(reviewer.acceptAll().absorbedBlocks).toEqual([
      { blockId: first.id, survivorBlockId: last.id },
    ]);
    const records = JSON.parse(
      JSON.stringify(reviewer.exportPendingSuggestions()),
    );
    expect(records).toMatchObject([
      {
        suggestionId: "pending",
        operation: { blockId: last.id },
        anchor: { blockId: last.id, paraId: last.id },
      },
    ]);
    const reopened = await FolioDocxReviewer.fromBuffer(
      await reviewer.toBuffer(),
      { session: "canonical" },
    );
    expect(reopened.loadPendingSuggestions(records)).toEqual([
      { status: "restaged", suggestionId: "pending" },
    ]);
    for (const proposal of [reviewer, reopened]) {
      expect(proposal.acceptSuggestion("pending")).toBe(true);
      expect(await savedTexts(proposal)).toEqual([
        "Opening clause.Signed in two copies.",
      ]);
    }
  });

  // On an absorbed paragraph or on the survivor itself.
  for (const [proposed, startOffset] of [
    ["Beta", 5],
    ["Gamma", 6],
  ] as const)
    test(`canonical bulk acceptance moves a range proposal on ${proposed} by the text absorbed before it`, async () => {
      const reviewer = await FolioDocxReviewer.fromBuffer(
        await documentOf(["Alpha one.", "Beta two.", "Gamma three."]),
        { session: "canonical" },
      );
      const [alpha, beta, gamma] = reviewer.getContent();
      if (!alpha || !beta || !gamma)
        throw new Error("fixture paragraphs missing");
      const word = proposed === "Beta" ? "two" : "three";
      applyOne(reviewer, "suggested", {
        id: "pending",
        type: "replaceRange",
        range: {
          type: "textRange",
          story: "main",
          blockId: proposed === "Beta" ? beta.id : gamma.id,
          startOffset,
          endOffset: startOffset + word.length,
          selectedTextHash: hashFolioAIBlockText(word),
        },
        replace: "2",
      });
      for (const block of [alpha, beta])
        applyOne(reviewer, "tracked-changes", {
          id: `merge-${block.id}`,
          type: "mergeBlockWithNext",
          blockId: block.id,
        });
      reviewer.acceptAll();
      const records = JSON.parse(
        JSON.stringify(reviewer.exportPendingSuggestions()),
      );
      const shifted = startOffset + (proposed === "Beta" ? 10 : 19);
      const range = { startOffset: shifted, endOffset: shifted + word.length };
      expect(records).toMatchObject([
        {
          operation: { range: { blockId: gamma.id, ...range } },
          anchor: { blockId: gamma.id, paraId: gamma.id, ...range },
        },
      ]);
      const reopened = await FolioDocxReviewer.fromBuffer(
        await reviewer.toBuffer(),
        { session: "canonical" },
      );
      expect(reopened.loadPendingSuggestions(records)).toEqual([
        { status: "restaged", suggestionId: "pending" },
      ]);
      for (const proposal of [reviewer, reopened]) {
        expect(proposal.acceptSuggestion("pending")).toBe(true);
        expect(await savedTexts(proposal)).toEqual([
          proposed === "Beta"
            ? "Alpha one.Beta 2.Gamma three."
            : "Alpha one.Beta two.Gamma 2.",
        ]);
      }
    });

  // As above: the canonical session records the suggested deletion beside the
  // tracked insertion, which stays as it was.
  test.each(SESSIONS)(
    "a suggested deletion leaves a tracked insertion pending (%s session)",
    async (session) => {
      const reviewer = await FolioDocxReviewer.fromBuffer(
        await buildDocument(),
        { author: "AI", session },
      );
      const first = reviewer.getContent()[0];
      if (!first) throw new Error("the fixture paragraph is missing");
      expect(
        reviewer.applyDocumentOperations(
          {
            version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
            mode: "tracked-changes",
            operations: [
              {
                id: "tracked",
                type: "insertAfterBlock",
                blockId: first.id,
                text: "Added.",
              },
            ],
          },
          { undefinedReferences: "refuse" },
        ).applied,
      ).toHaveLength(1);
      const added = reviewer.getContent().find(({ text }) => text === "Added.");
      if (!added) throw new Error("the tracked insertion is missing");
      const deletion = reviewer.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode: "suggested",
          operations: [
            { id: "delete", type: "deleteBlock", blockId: added.id },
          ],
        },
        { undefinedReferences: "refuse" },
      );
      if (session === "legacy") {
        expect(deletion.applied).toEqual([]);
        expect(deletion.skipped).toEqual([
          { id: "delete", reason: "unsupportedMode" },
        ]);
      } else {
        expect(deletion.applied.map(({ id }) => id)).toEqual(["delete"]);
        expect(deletion.skipped).toEqual([]);
      }
      reviewer.acceptAll();
      expect(await savedTexts(reviewer)).toEqual([
        "First clause.",
        "Added.",
        "Signed in two copies.",
      ]);
    },
  );

  // Beside a paragraph a tracked deletion takes whole, the legacy engine
  // refuses a suggested block; the canonical session puts it there, as the
  // `pendingDeletion` refusal's recovery says ("insert a new block next to
  // it").
  test.each(["legacy", "canonical"] as const)(
    "a suggestion anchored to a tracked deletion (%s session)",
    async (session) => {
      const writer = await FolioDocxReviewer.fromBuffer(await buildDocument(), {
        author: "AI",
        session,
      });
      const first = writer.getContent()[0];
      if (!first) throw new Error("fixture paragraph missing");
      expect(
        writer.applyDocumentOperations(
          {
            version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
            mode: "tracked-changes",
            operations: [
              { id: "tracked", type: "deleteBlock", blockId: first.id },
            ],
          },
          { undefinedReferences: "refuse" },
        ).applied,
      ).toHaveLength(1);
      const bytes = await writer.toBuffer();
      for (const type of ["insertAfterBlock", "insertBeforeBlock"] as const) {
        const baseline = await FolioDocxReviewer.fromBuffer(bytes, {
          author: "AI",
          session,
        });
        const staged = await FolioDocxReviewer.fromBuffer(bytes, {
          author: "AI",
          session,
        });
        const blank = staged.getContent().find(({ text }) => text === "");
        if (!blank) throw new Error("deleted paragraph missing");
        const result = staged.applyDocumentOperations(
          {
            version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
            mode: "suggested",
            operations: [
              { id: "pending", type, blockId: blank.id, text: "Added." },
            ],
          },
          { undefinedReferences: "refuse" },
        );
        if (session === "canonical") {
          expect(result.skipped).toEqual([]);
          expect(result.applied.map(({ id }) => id)).toEqual(["pending"]);
          staged.acceptAll();
          expect(staged.getContent().map(({ text }) => text)).toContain(
            "Added.",
          );
          continue;
        }
        expect(result.applied).toEqual([]);
        expect(result.skipped).toEqual([
          { id: "pending", reason: "pendingDeletion" },
        ]);
        const beforeBulk = await FolioDocxReviewer.fromBuffer(
          await staged.toBuffer(),
          { session },
        );
        expect(beforeBulk.getContent()).toEqual(baseline.getContent());
        baseline.acceptAll();
        staged.acceptAll();
        expect(
          (
            await FolioDocxReviewer.fromBuffer(await staged.toBuffer(), {
              session,
            })
          ).getContent(),
        ).toEqual(
          (
            await FolioDocxReviewer.fromBuffer(await baseline.toBuffer(), {
              session,
            })
          ).getContent(),
        );
      }
    },
  );

  test("suggested replacement and deletion preserve the saved heading", async () => {
    const doc = fromMarkdown(
      "# Delivery Terms\n\nThe following apply to every order.\n\n- Goods are packed securely\n- Goods are insured in transit\n\nPayment happens in stages.\n\n1. Deposit on signature\n2. Balance on delivery\n3. Retention after inspection\n\nClosing remarks.",
    );
    const { docx } = await ensureParaIds(new Uint8Array(await createDocx(doc)));
    const reviewer = await FolioDocxReviewer.fromBuffer(
      docx.buffer.slice(
        docx.byteOffset,
        docx.byteOffset + docx.byteLength,
      ) as ArrayBuffer,
      { author: "AI" },
    );
    const block = (prefix: string) =>
      reviewer.getContent().find(({ text }) => text.startsWith(prefix));
    const opening = block("The following");
    const payment = block("Payment happens");
    const insured = block("Goods are insured");
    if (!opening || !payment || !insured)
      throw new Error("fixture block missing");
    const tracked = reviewer.applyDocumentOperations(
      {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "tracked-changes",
        operations: [
          {
            id: "t1",
            type: "replaceInBlock",
            blockId: opening.id,
            find: "every",
            replace: "each",
          },
          {
            id: "t2",
            type: "insertAfterBlock",
            blockId: payment.id,
            text: "Stages are invoiced separately.",
          },
          { id: "t3", type: "deleteBlock", blockId: insured.id },
        ],
      },
      { undefinedReferences: "refuse" },
    );
    expect(tracked.applied).toHaveLength(3);
    const trackedReopened = await FolioDocxReviewer.fromBuffer(
      await reviewer.toBuffer(),
      {
        author: "AI",
      },
    );
    const baseline = await FolioDocxReviewer.fromBuffer(
      await reviewer.toBuffer(),
      {
        author: "AI",
      },
    );
    baseline.acceptAll();
    const baselineContent = (
      await FolioDocxReviewer.fromBuffer(await baseline.toBuffer())
    ).getContent();
    const heading = trackedReopened.getContent()[0];
    if (!heading) throw new Error("heading missing");
    const result = trackedReopened.applyDocumentOperations(
      {
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
      },
      { undefinedReferences: "refuse" },
    );
    expect(result.applied.map(({ id }) => id)).toEqual(["replace"]);
    expect(result.skipped.map(({ id }) => id)).toEqual(["delete"]);
    trackedReopened.acceptAll();
    const reopened = await FolioDocxReviewer.fromBuffer(
      await trackedReopened.toBuffer(),
    );
    expect(reopened.getContent()).toEqual(baselineContent);
  });

  test("acceptChange on a suggested paragraph insert keeps the paragraph through the save", async () => {
    const reviewer = await suggest();
    const insertion = reviewer
      .getChanges()
      .find(
        ({ type, text }) =>
          type === "insertion" && text === "Suggested clause.",
      );
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
      .find(
        ({ type, text }) =>
          type === "insertion" && text === "Suggested clause.",
      );
    if (!insertion) {
      throw new Error("the suggested insertion is not listed");
    }
    expect(reviewer.rejectChange(insertion)).toBe(true);
    expect(reviewer.getContent().map(({ text }) => text)).not.toContain("");
    expect(reviewer.getContent().map(({ text }) => text)).not.toContain(
      "Suggested clause.",
    );
  });

  test("rejectAll removes every suggestion", async () => {
    const reviewer = await suggest();
    reviewer.rejectAll();
    expect(reviewer.getContent().map(({ text }) => text)).toEqual([
      "First clause.",
      "Signed in two copies.",
    ]);
    expect(await savedTexts(reviewer)).toEqual([
      "First clause.",
      "Signed in two copies.",
    ]);
  });

  for (const accept of ["acceptChange", "acceptSuggestion"] as const) {
    test(`a suggested paragraph after a break pending deletion persists once accepted (${accept})`, async () => {
      const reviewer = await FolioDocxReviewer.fromBuffer(
        await buildDocument(),
        { author: "AI" },
      );
      const first = reviewer
        .getContent()
        .find(({ text }) => text === "First clause.");
      if (!first) throw new Error("the fixture paragraph is missing");
      // The join leaves the paragraph's break pending deletion; the suggested
      // paragraph inserted after it then carries that break.
      const joined = reviewer.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode: "tracked-changes",
          operations: [
            { id: "join", type: "mergeBlockWithNext", blockId: first.id },
          ],
        },
        { undefinedReferences: "refuse" },
      );
      expect(joined.applied).toHaveLength(1);
      const inserted = reviewer.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode: "suggested",
          operations: [
            {
              id: "heading",
              type: "insertAfterBlock",
              blockId: first.id,
              text: "Added heading.",
              styleId: "Heading2",
            },
          ],
        },
        { undefinedReferences: "refuse" },
      );
      expect(inserted.applied).toHaveLength(1);
      if (accept === "acceptChange") {
        const insertion = reviewer
          .getChanges()
          .find(
            ({ type, text }) =>
              type === "insertion" && text === "Added heading.",
          );
        if (!insertion) throw new Error("the suggested insertion is missing");
        expect(reviewer.acceptChange(insertion)).toBe(true);
      } else {
        expect(reviewer.acceptSuggestion("heading")).toBe(true);
      }
      const live = reviewer.getContent().map(({ text }) => text);
      expect(live).toEqual([
        "First clause.",
        "Added heading.",
        "Signed in two copies.",
      ]);
      expect(await savedTexts(reviewer)).toEqual(live);
      // The join stays pending, now ending at the accepted paragraph.
      expect(reviewer.getChanges().map(({ type }) => type)).toContain(
        "paragraphMarkDeleted",
      );
    });
  }
});

// Reject All rejects proposals in the commit that takes a pending inserted
// paragraph out again: a proposal on that paragraph ends with it and has no
// survivor to follow. The registry follows survivors on the canonical session
// only; the legacy one restages proposals instead.
describe("rejecting all with a proposal on a pending inserted paragraph", () => {
  const proposals = (blockId: string): FolioDocumentOperation[] => [
    {
      id: "replace",
      type: "replaceInBlock",
      blockId,
      find: "are",
      replace: "get",
    },
    {
      id: "range",
      type: "replaceRange",
      range: {
        type: "textRange",
        story: "main",
        blockId,
        startOffset: 0,
        endOffset: 6,
        selectedTextHash: hashFolioAIBlockText("Stages"),
      },
      replace: "Parts",
    },
    {
      id: "rewrite",
      type: "replaceBlock",
      blockId,
      text: "Stages are billed.",
    },
    {
      id: "comment",
      type: "commentOnBlock",
      blockId,
      quote: "are",
      comment: { text: "Check." },
    },
    {
      id: "style",
      type: "setBlockParagraphProperties",
      blockId,
      properties: { alignment: "center" },
    },
    {
      id: "before",
      type: "insertBeforeBlock",
      blockId,
      text: "Added clause.",
    },
  ];
  for (const { id } of proposals("x"))
    test(`${id} (canonical session)`, async () => {
      const source = ["Payment happens in stages.", "Deposit on signature."];
      const reviewer = await FolioDocxReviewer.fromBuffer(
        await documentOf(source),
        { author: "Editor", session: "canonical" },
      );
      const [payment] = reviewer.getContent();
      if (!payment) throw new Error("fixture paragraph missing");
      applyOne(reviewer, "tracked-changes", {
        id: "insert",
        type: "insertAfterBlock",
        blockId: payment.id,
        text: "Stages are invoiced separately.",
      });
      const inserted = reviewer
        .getContent()
        .find(({ text }) => text === "Stages are invoiced separately.");
      if (!inserted) throw new Error("inserted paragraph missing");
      const proposal = proposals(inserted.id).find(
        (candidate) => candidate.id === id,
      );
      if (!proposal) throw new Error("proposal missing");
      applyOne(reviewer, "suggested", proposal);
      reviewer.rejectAll();
      expect(reviewer.getContent().map(({ text }) => text)).toEqual(source);
      expect(reviewer.exportPendingSuggestions()).toEqual([]);
      expect(await savedTexts(reviewer)).toEqual(source);
    });
});
