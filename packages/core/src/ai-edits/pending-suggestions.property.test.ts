/**
 * The old suggestion/save property checked that proposed text vanished from a
 * DOCX, but had no host-store fixture or rehydration oracle. This property
 * adds that missing export, JSON, reopen, and load sequence; targeted stale
 * cases exercise anchor identity rather than only example text matching.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";
import { createDocx } from "../docx/rezip";
import { parseDocx } from "../docx/parser";
import { repackDocx } from "../docx/rezip";
import { paragraph } from "../docx/server/build";
import { schema } from "../prosemirror/schema";
import type { HeaderFooter } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { createFolioAIEditSnapshot, createFolioAITextRangeHandle } from "./snapshot";
import { FolioDocxReviewer } from "./headless";
import { FolioPendingSuggestionRegistry } from "./pending-suggestions";

const FIRST_ID = "10000001";
const SECOND_ID = "10000002";
const FIRST_TEXT = "First clause remains.";

const documentBytes = async (secondText: string, includeSecond = true): Promise<ArrayBuffer> => {
  const document = createEmptyDocument();
  document.package.document.content = [
    { ...paragraph(FIRST_TEXT), paraId: FIRST_ID },
    ...(includeSecond ? [{ ...paragraph(secondText), paraId: SECOND_ID }] : []),
  ];
  return createDocx(document);
};

const projection = (reviewer: FolioDocxReviewer) =>
  reviewer.snapshot().blocks.map(({ id, kind, text }) => ({ id, kind, text }));

const stage = async (text: string, replacement: string) => {
  const reviewer = await FolioDocxReviewer.fromBuffer(await documentBytes(text));
  const result = reviewer.applyOperations(
    [
      {
        id: "proposal-1",
        type: "replaceInBlock",
        blockId: SECOND_ID,
        find: text,
        replace: replacement,
      },
    ],
    { mode: "suggested" },
  );
  expect(result.applied).toHaveLength(1);
  const records = JSON.parse(JSON.stringify(reviewer.exportPendingSuggestions()));
  expect(records).toHaveLength(1);
  return { reviewer, records };
};

describe("host-persisted pending suggestions", () => {
  test(
    "export, JSON save, reopen, and load restore the live projection",
    async () => {
      await fc.assert(
        fc.asyncProperty(
          fc.stringMatching(/^[A-Za-z]{3,12}$/u),
          fc.stringMatching(/^[A-Za-z]{3,12}$/u),
          async (original, replacement) => {
            fc.pre(original !== replacement);
            const text = `Clause ${original}.`;
            const { reviewer, records } = await stage(text, `Clause ${replacement}.`);
            const before = projection(reviewer);
            const reopened = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
            expect(projection(reopened)).not.toEqual(before);
            expect(reopened.loadPendingSuggestions(records)).toEqual([
              { status: "restaged", suggestionId: "proposal-1" },
            ]);
            expect(projection(reopened)).toEqual(before);
            expect(reopened.exportPendingSuggestions()).toEqual(records);
            expect(reopened.loadPendingSuggestions(records)).toEqual([
              { status: "restaged", suggestionId: "proposal-1" },
            ]);
            expect(projection(reopened)).toEqual(before);
            expect(reopened.acceptAll()).toBe(0);
            expect(reopened.exportPendingSuggestions()).toEqual(records);
          },
        ),
        propertyConfig({ numRuns: 12, seed: 174496237 }),
      );
    },
    propertyTestTimeout(30_000),
  );

  test("changed text and deleted anchors stay inert", async () => {
    const { records } = await stage("Second clause.", "Revised clause.");
    const changed = await FolioDocxReviewer.fromBuffer(await documentBytes("Changed clause."));
    const changedBefore = projection(changed);
    expect(changed.loadPendingSuggestions(records)).toEqual([
      { status: "stale", suggestionId: "proposal-1", reason: "textChanged" },
    ]);
    expect(projection(changed)).toEqual(changedBefore);
    expect(changed.exportPendingSuggestions()).toEqual([]);

    const deleted = await FolioDocxReviewer.fromBuffer(await documentBytes("", false));
    const deletedBefore = projection(deleted);
    expect(deleted.loadPendingSuggestions(records)).toEqual([
      { status: "stale", suggestionId: "proposal-1", reason: "missingAnchor" },
    ]);
    expect(projection(deleted)).toEqual(deletedBefore);
  });

  test("range records retain offsets and the selected-text hash", async () => {
    const text = "Second clause.";
    const reviewer = await FolioDocxReviewer.fromBuffer(await documentBytes(text));
    const range = createFolioAITextRangeHandle({
      blockId: SECOND_ID,
      text,
      startOffset: 0,
      endOffset: 6,
    });
    const result = reviewer.applyOperations(
      [{ id: "range-proposal", type: "replaceRange", range, replace: "Revised" }],
      { mode: "suggested" },
    );
    expect(result.applied).toHaveLength(1);
    const record = reviewer.exportPendingSuggestions().at(0);
    expect(record?.anchor).toMatchObject({
      selectedTextHash: range.selectedTextHash,
      startOffset: range.startOffset,
      endOffset: range.endOffset,
    });
  });

  test("loading retains the original proposal author", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await documentBytes("Second clause."), {
      author: "Origin Author",
    });
    const result = reviewer.applyOperations(
      [
        {
          id: "authored",
          type: "replaceInBlock",
          blockId: SECOND_ID,
          find: "Second",
          replace: "Revised",
        },
      ],
      { mode: "suggested" },
    );
    expect(result.applied).toHaveLength(1);
    const records = JSON.parse(JSON.stringify(reviewer.exportPendingSuggestions()));
    const reopened = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer(), {
      author: "Review Author",
    });
    expect(reopened.loadPendingSuggestions(records)).toEqual([
      { status: "restaged", suggestionId: "authored" },
    ]);
    expect(new Set(reopened.getChanges().map(({ author }) => author))).toEqual(
      new Set(["Origin Author"]),
    );
  });

  test("sequential proposals in one paragraph replay in order", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await documentBytes("Second clause."));
    for (const [id, find, replace] of [
      ["first-proposal", "Second", "Revised"],
      ["second-proposal", "clause", "section"],
    ]) {
      const result = reviewer.applyOperations(
        [{ id, type: "replaceInBlock", blockId: SECOND_ID, find, replace }],
        { mode: "suggested" },
      );
      expect(result.applied).toHaveLength(1);
    }
    const records = JSON.parse(JSON.stringify(reviewer.exportPendingSuggestions()));
    expect(records).toHaveLength(2);
    const before = projection(reviewer);
    const reopened = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
    expect(reopened.loadPendingSuggestions(records)).toEqual([
      { status: "restaged", suggestionId: "first-proposal" },
      { status: "restaged", suggestionId: "second-proposal" },
    ]);
    expect(projection(reopened)).toEqual(before);
    const loadedInBatches = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
    expect(loadedInBatches.loadPendingSuggestions(records.slice(0, 1))).toEqual([
      { status: "restaged", suggestionId: "first-proposal" },
    ]);
    expect(loadedInBatches.loadPendingSuggestions(records.slice(1))).toEqual([
      { status: "restaged", suggestionId: "second-proposal" },
    ]);
    expect(projection(loadedInBatches)).toEqual(before);
  });

  test("replay records the comment id allocated by the host", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await documentBytes("Second clause."));
    expect(
      reviewer.applyOperations(
        [
          {
            id: "commented-proposal",
            type: "replaceInBlock",
            blockId: SECOND_ID,
            find: "Second",
            replace: "Revised",
            comment: { text: "Review this edit." },
          },
        ],
        { mode: "suggested" },
      ).applied,
    ).toHaveLength(1);
    const records = reviewer.exportPendingSuggestions();
    expect(records.at(0)?.commentId).toBeDefined();
    const source = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
    const snapshot = source.snapshot();
    const registry = new FolioPendingSuggestionRegistry();
    let active = false;
    expect(
      registry.loadPendingSuggestions({
        records,
        snapshotForStory: () => snapshot,
        sourceSnapshotForStory: () => snapshot,
        apply: (record) => {
          active = true;
          return {
            applied: [
              { id: record.operation.id, suggestionId: record.suggestionId, commentId: 999 },
            ],
            skipped: [],
          };
        },
        activeSuggestionIds: () => new Set(active ? ["commented-proposal"] : []),
      }),
    ).toEqual([{ status: "restaged", suggestionId: "commented-proposal" }]);
    const replayed = registry.exportPendingSuggestions({
      activeSuggestionIds: new Set(["commented-proposal"]),
      snapshotForStory: () => snapshot,
    });
    expect(replayed.at(0)?.commentId).toBe(999);
  });

  test("a formatting-only document change invalidates a saved proposal", async () => {
    const { records } = await stage("Second clause.", "Revised clause.");
    const changed = await FolioDocxReviewer.fromBuffer(await documentBytes("Second clause."));
    const range = createFolioAITextRangeHandle({
      blockId: SECOND_ID,
      text: "Second clause.",
      startOffset: 0,
      endOffset: 6,
    });
    expect(
      changed.applyOperations(
        [{ id: "direct-format", type: "formatRange", range, formatting: { bold: true } }],
        { mode: "direct" },
      ).applied,
    ).toHaveLength(1);
    expect(changed.loadPendingSuggestions(records)).toEqual([
      { status: "stale", suggestionId: "proposal-1", reason: "documentChanged" },
    ]);
  });

  test("a formatted baseline keeps the same fingerprint across a save", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await documentBytes("Second clause."));
    const range = createFolioAITextRangeHandle({
      blockId: SECOND_ID,
      text: "Second clause.",
      startOffset: 0,
      endOffset: 6,
    });
    expect(
      reviewer.applyOperations(
        [{ id: "baseline-format", type: "formatRange", range, formatting: { bold: true } }],
        { mode: "direct" },
      ).applied,
    ).toHaveLength(1);
    expect(
      reviewer.applyOperations(
        [
          {
            id: "formatted-proposal",
            type: "replaceInBlock",
            blockId: SECOND_ID,
            find: "clause",
            replace: "section",
          },
        ],
        { mode: "suggested" },
      ).applied,
    ).toHaveLength(1);
    const records = JSON.parse(JSON.stringify(reviewer.exportPendingSuggestions()));
    const reopened = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
    expect(reopened.loadPendingSuggestions(records)).toEqual([
      { status: "restaged", suggestionId: "formatted-proposal" },
    ]);
  });

  test("an unrelated direct edit after staging refreshes the saved story fingerprint", async () => {
    const { reviewer } = await stage("Second clause.", "Revised clause.");
    expect(
      reviewer.applyOperations(
        [
          {
            id: "later-edit",
            type: "replaceInBlock",
            blockId: FIRST_ID,
            find: "First",
            replace: "Opening",
          },
        ],
        { mode: "direct" },
      ).applied,
    ).toHaveLength(1);
    const records = JSON.parse(JSON.stringify(reviewer.exportPendingSuggestions()));
    const reopened = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
    expect(reopened.loadPendingSuggestions(records)).toEqual([
      { status: "restaged", suggestionId: "proposal-1" },
    ]);
  });

  test("secondary-story proposals retain their story on reload", async () => {
    const relationshipId = "rId_pending_header";
    const story = { type: "header", relationshipId } as const;
    const document = await parseDocx(await documentBytes("Second clause."), {
      detectVariables: false,
      preloadFonts: false,
    });
    const header: HeaderFooter = {
      type: "header",
      hdrFtrType: "default",
      content: [{ ...paragraph("Header clause."), paraId: "10000003" }],
    };
    document.package.headers = new Map([[relationshipId, header]]);
    document.package.document.finalSectionProperties = {
      ...document.package.document.finalSectionProperties,
      headerReferences: [{ type: "default", rId: relationshipId }],
    };
    const bytes = await repackDocx(document, { updateModifiedDate: false });
    const reviewer = await FolioDocxReviewer.fromBuffer(bytes);
    const snapshot = reviewer.snapshotStory(story);
    expect(snapshot).not.toBeNull();
    const result = reviewer.applyDocumentOperationsToStory({
      story,
      batch: {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "suggested",
        operations: [
          {
            id: "header-proposal",
            type: "replaceInBlock",
            blockId: "10000003",
            find: "Header",
            replace: "Revised",
          },
        ],
      },
    });
    expect(result.applied).toHaveLength(1);
    const records = JSON.parse(JSON.stringify(reviewer.exportPendingSuggestions()));
    expect(records.at(0)?.story).toEqual(story);
    const live = reviewer.snapshotStory(story)?.blocks.map(({ text }) => text);
    const reopened = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
    expect(reopened.loadPendingSuggestions(records)).toEqual([
      { status: "restaged", suggestionId: "header-proposal" },
    ]);
    expect(reopened.snapshotStory(story)?.blocks.map(({ text }) => text)).toEqual(live);
  });

  test("a duplicated paragraph id cannot redirect a saved proposal", async () => {
    const { records } = await stage("Second clause.", "Revised clause.");
    const duplicateDoc = schema.node("doc", null, [
      schema.node("paragraph", { paraId: SECOND_ID }, [schema.text("Second clause.")]),
      schema.node("paragraph", { paraId: SECOND_ID }, [schema.text("Second clause.")]),
    ]);
    const snapshot = createFolioAIEditSnapshot(duplicateDoc);
    const registry = new FolioPendingSuggestionRegistry();
    let applied = false;
    expect(
      registry.loadPendingSuggestions({
        records,
        snapshotForStory: () => snapshot,
        sourceSnapshotForStory: () => snapshot,
        apply: () => {
          applied = true;
          return { applied: [], skipped: [] };
        },
        activeSuggestionIds: () => new Set(),
      }),
    ).toEqual([{ status: "stale", suggestionId: "proposal-1", reason: "ambiguousAnchor" }]);
    expect(applied).toBe(false);
  });

  test("unsupported record versions stay inert", async () => {
    const { records } = await stage("Second clause.", "Revised clause.");
    const target = await FolioDocxReviewer.fromBuffer(await documentBytes("Second clause."));
    const record = records.at(0);
    expect(target.loadPendingSuggestions([{ ...record, version: 2 }])).toEqual([
      { status: "stale", suggestionId: "proposal-1", reason: "unsupportedVersion" },
    ]);
    expect(target.exportPendingSuggestions()).toEqual([]);
  });

  test("a loaded suggestion id cannot mask a different operation", async () => {
    const { records } = await stage("Second clause.", "Revised clause.");
    const reviewer = await FolioDocxReviewer.fromBuffer(await documentBytes("Second clause."));
    expect(reviewer.loadPendingSuggestions(records)).toEqual([
      { status: "restaged", suggestionId: "proposal-1" },
    ]);
    const changed = structuredClone(records);
    changed[0].operation.replace = "Different clause.";
    expect(reviewer.loadPendingSuggestions(changed)).toEqual([
      { status: "stale", suggestionId: "proposal-1", reason: "invalidRecord" },
    ]);
    expect(reviewer.exportPendingSuggestions()).toEqual(records);
  });
});
