/**
 * `createEditorRefBridge` tests against a fake `FolioAgentEditorRefLike`,
 * covering the read-surface upgrade: `getChanges` mapping via
 * `getTrackedChanges`, comment-anchor merging via `getCommentAnchors`, and
 * the capability-conditional `getSelectionText` / `getPageText` bridge
 * members (present only when the ref implements the underlying method).
 */

import { describe, expect, test } from "bun:test";

import type { FolioAIEditApplyResult, FolioAIEditSnapshot } from "@stll/folio-core/server";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "@stll/folio-core/server";
import type { FolioCommentAnchor, FolioReviewChange } from "@stll/folio-core/ai-edits";
import type { Comment } from "@stll/folio-core/types/content";

import { createEditorRefBridge, type FolioAgentEditorRefLike } from "./editor-ref";

const EMPTY_SNAPSHOT: FolioAIEditSnapshot = { blocks: [], anchors: {} };
const EMPTY_APPLY_RESULT: FolioAIEditApplyResult = { applied: [], skipped: [] };

/** Base fake ref implementing only the required `FolioAgentEditorRefLike` members (an "older ref"). */
const baseRef = (): FolioAgentEditorRefLike => ({
  createAIEditSnapshot: () => EMPTY_SNAPSHOT,
  applyAIEditOperations: () => EMPTY_APPLY_RESULT,
  scrollToBlock: () => true,
  getTotalPages: () => 3,
});

describe("createEditorRefBridge: document operations", () => {
  test("delegates a versioned batch to a current editor ref", () => {
    let receivedVersion: number | undefined;
    const ref: FolioAgentEditorRefLike = {
      ...baseRef(),
      applyDocumentOperations: ({ batch }) => {
        receivedVersion = batch.version;
        return { version: batch.version, status: "committed", applied: [], skipped: [] };
      },
    };
    const bridge = createEditorRefBridge({
      ref,
      author: "AI",
      getComments: () => [],
      setComments: () => {},
    });

    const result = bridge.applyDocumentOperations({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      operations: [],
    });

    expect(receivedVersion).toBe(FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION);
    expect(result.version).toBe(FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION);
  });

  test("preserves the versioned result when adapting an older editor ref", () => {
    const bridge = createEditorRefBridge({
      ref: baseRef(),
      author: "AI",
      getComments: () => [],
      setComments: () => {},
    });

    const result = bridge.applyDocumentOperations({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      operations: [],
    });

    expect(result).toEqual({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      status: "committed",
      ...EMPTY_APPLY_RESULT,
    });
  });

  test("rejects atomic batches without mutating through an older editor ref", () => {
    let applyCalls = 0;
    const ref: FolioAgentEditorRefLike = {
      ...baseRef(),
      applyAIEditOperations: () => {
        applyCalls += 1;
        return EMPTY_APPLY_RESULT;
      },
    };
    const bridge = createEditorRefBridge({
      ref,
      author: "AI",
      getComments: () => [],
      setComments: () => {},
    });

    const result = bridge.applyDocumentOperations({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      atomic: true,
      operations: [{ id: "op-1", type: "deleteBlock", blockId: "para-1" }],
    });

    expect(result).toEqual({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      status: "rejected",
      applied: [],
      skipped: [{ id: "op-1", reason: "unsupportedMode" }],
    });
    expect(applyCalls).toBe(0);
  });

  test("reports dry runs as unsupported without mutating through an older editor ref", () => {
    let applyCalls = 0;
    const ref: FolioAgentEditorRefLike = {
      ...baseRef(),
      applyAIEditOperations: () => {
        applyCalls += 1;
        return EMPTY_APPLY_RESULT;
      },
    };
    const bridge = createEditorRefBridge({
      ref,
      author: "AI",
      getComments: () => [],
      setComments: () => {},
    });

    const result = bridge.applyDocumentOperations({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      dryRun: true,
      operations: [{ id: "op-1", type: "deleteBlock", blockId: "para-1" }],
    });

    expect(result).toEqual({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      status: "previewed",
      applied: [],
      skipped: [{ id: "op-1", reason: "unsupportedMode" }],
    });
    expect(applyCalls).toBe(0);
  });

  test("rejects unsupported serialized versions before using a legacy ref", () => {
    const bridge = createEditorRefBridge({
      ref: baseRef(),
      author: "AI",
      getComments: () => [],
      setComments: () => {},
    });

    expect(() =>
      Reflect.apply(bridge.applyDocumentOperations, bridge, [{ version: 2, operations: [] }]),
    ).toThrow("Unsupported document operation contract version.");
  });
});

const makeComment = (id: number, parentId?: number): Comment => ({
  id,
  author: "Tester",
  date: "2026-05-15T00:00:00Z",
  content: [
    {
      type: "paragraph",
      formatting: {},
      content: [
        { type: "run", formatting: {}, content: [{ type: "text", text: `comment-${id}` }] },
      ],
    },
  ],
  ...(parentId !== undefined ? { parentId } : {}),
});

describe("createEditorRefBridge: getChanges", () => {
  test("maps ref.getTrackedChanges output the same way the reviewer bridge maps FolioReviewChange", () => {
    const change: FolioReviewChange = {
      id: 7,
      type: "insertion",
      author: "AI Reviewer",
      date: "2026-05-15T00:00:00Z",
      text: "inserted text",
      blockId: "block-1",
    };
    const ref: FolioAgentEditorRefLike = {
      ...baseRef(),
      getTrackedChanges: () => [change],
    };
    const bridge = createEditorRefBridge({
      ref,
      author: "AI",
      getComments: () => [],
      setComments: () => {},
    });

    expect(bridge.getChanges()).toEqual([
      {
        id: "7",
        type: "insertion",
        author: "AI Reviewer",
        text: "inserted text",
        blockId: "block-1",
      },
    ]);
  });

  test("returns [] when the ref does not implement getTrackedChanges (older ref)", () => {
    const bridge = createEditorRefBridge({
      ref: baseRef(),
      author: "AI",
      getComments: () => [],
      setComments: () => {},
    });

    expect(bridge.getChanges()).toEqual([]);
  });
});

describe("createEditorRefBridge: getComments", () => {
  test("merges anchor blockId/quote into the matching host-state comment by commentId", () => {
    const anchors: FolioCommentAnchor[] = [
      { commentId: 1, blockId: "block-a", quote: "quoted text" },
      { commentId: 2, blockId: "block-b", quote: "other quote" },
    ];
    const ref: FolioAgentEditorRefLike = {
      ...baseRef(),
      getCommentAnchors: () => anchors,
    };
    const comments = [makeComment(1), makeComment(2)];
    const bridge = createEditorRefBridge({
      ref,
      author: "AI",
      getComments: () => comments,
      setComments: () => {},
    });

    const result = bridge.getComments();
    expect(result).toHaveLength(2);
    expect(result.find((c) => c.id === "1")).toMatchObject({
      blockId: "block-a",
      quote: "quoted text",
    });
    expect(result.find((c) => c.id === "2")).toMatchObject({
      blockId: "block-b",
      quote: "other quote",
    });
  });

  test("leaves blockId null and quote empty when the ref does not implement getCommentAnchors (older ref)", () => {
    const bridge = createEditorRefBridge({
      ref: baseRef(),
      author: "AI",
      getComments: () => [makeComment(1)],
      setComments: () => {},
    });

    const result = bridge.getComments();
    expect(result).toHaveLength(1);
    expect(result[0]).toMatchObject({ blockId: null, quote: "" });
  });

  test("a comment with no matching anchor entry also falls back to null/empty", () => {
    const ref: FolioAgentEditorRefLike = {
      ...baseRef(),
      getCommentAnchors: () => [{ commentId: 999, blockId: "block-z", quote: "unrelated" }],
    };
    const bridge = createEditorRefBridge({
      ref,
      author: "AI",
      getComments: () => [makeComment(1)],
      setComments: () => {},
    });

    const result = bridge.getComments();
    expect(result[0]).toMatchObject({ blockId: null, quote: "" });
  });
});

describe("createEditorRefBridge: capability-conditional members", () => {
  test("getSelectionText is present and delegates when the ref implements it", () => {
    const ref: FolioAgentEditorRefLike = {
      ...baseRef(),
      getSelectionText: () => "selected text",
    };
    const bridge = createEditorRefBridge({
      ref,
      author: "AI",
      getComments: () => [],
      setComments: () => {},
    });

    expect(bridge.getSelectionText).toBeDefined();
    expect(bridge.getSelectionText?.()).toBe("selected text");
  });

  test("getSelectionText is absent (not a no-op) when the ref does not implement it", () => {
    const bridge = createEditorRefBridge({
      ref: baseRef(),
      author: "AI",
      getComments: () => [],
      setComments: () => {},
    });

    expect(bridge.getSelectionText).toBeUndefined();
  });

  test('getPageText is present, delegates, and maps a null page (in-range, layout not yet computed) to ""', () => {
    const ref: FolioAgentEditorRefLike = {
      ...baseRef(),
      getPageText: (page) => (page === 1 ? "page one text" : null),
    };
    const bridge = createEditorRefBridge({
      ref,
      author: "AI",
      getComments: () => [],
      setComments: () => {},
    });

    expect(bridge.getPageText).toBeDefined();
    expect(bridge.getPageText?.(1)).toBe("page one text");
    expect(bridge.getPageText?.(2)).toBe("");
  });

  test("getPageText is absent (not a no-op) when the ref does not implement it", () => {
    const bridge = createEditorRefBridge({
      ref: baseRef(),
      author: "AI",
      getComments: () => [],
      setComments: () => {},
    });

    expect(bridge.getPageText).toBeUndefined();
  });

  test("getPageCount always delegates to ref.getTotalPages (required member, unaffected by the optional ones)", () => {
    const bridge = createEditorRefBridge({
      ref: baseRef(),
      author: "AI",
      getComments: () => [],
      setComments: () => {},
    });

    expect(bridge.getPageCount?.()).toBe(3);
  });
});
