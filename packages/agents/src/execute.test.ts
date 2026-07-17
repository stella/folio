/**
 * Tool-execution tests: argument validation (unknown tool, missing/wrong-type
 * args, unsupported capability), then a full happy path against a real
 * `FolioDocxReviewer` built from a docx fixture, exercising the reviewer
 * bridge exactly as a host app would.
 */

import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";

import { FolioDocxReviewer } from "@stll/folio-core/server";

import { createReviewerBridge } from "./bridges/reviewer";
import type { FolioAgentBridge } from "./bridge";
import { executeFolioToolCall } from "./execute";
import { FOLIO_AGENT_TOOL_NAMES } from "./types";
import type {
  FolioAgentApplyOperationsSummary,
  FolioAgentBlock,
  FolioAgentComment,
  FolioAgentFindTextResult,
  FolioAgentDocumentOutline,
  FolioAgentSectionRead,
  FolioAgentStoryFindTextResult,
  FolioToolCallResult,
} from "./types";

// Reuses the same corpus fixture `packages/core/src/ai-edits/headless.test.ts`
// builds its reviewer round-trip tests against.
const FIXTURE = path.join(
  import.meta.dir,
  "../../core/src/docx/__tests__/__fixtures__/corpus/authored-empty-paragraph.docx",
);

const readFixture = (): ArrayBuffer => {
  const bytes = readFileSync(FIXTURE);
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
};

const expectOk = (result: FolioToolCallResult): unknown => {
  if (!result.ok) {
    throw new Error(`expected ok:true, got error: ${result.error}`);
  }
  return result.result;
};

const expectError = (result: FolioToolCallResult): string => {
  if (result.ok) {
    throw new Error("expected ok:false, got a result");
  }
  return result.error;
};

describe("executeFolioToolCall: argument validation", () => {
  test("unknown tool name is rejected with the valid tool list", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(readFixture());
    const bridge = createReviewerBridge(reviewer);

    const result = executeFolioToolCall("not_a_real_tool", {}, bridge);
    const error = expectError(result);
    expect(error).toContain("not_a_real_tool");
    expect(error).toContain("read_document");
  });

  test("missing required argument is rejected", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(readFixture());
    const bridge = createReviewerBridge(reviewer);

    const result = executeFolioToolCall(FOLIO_AGENT_TOOL_NAMES.findText, {}, bridge);
    expect(expectError(result)).toContain("query");
  });

  test("wrong-type argument is rejected", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(readFixture());
    const bridge = createReviewerBridge(reviewer);

    const result = executeFolioToolCall(
      FOLIO_AGENT_TOOL_NAMES.findText,
      { query: "Heading", matchCase: "yes" },
      bridge,
    );
    expect(expectError(result)).toContain("matchCase");
  });

  test("unsupported capability on the headless reviewer bridge reports a plain-language error", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(readFixture());
    const bridge = createReviewerBridge(reviewer);

    const result = executeFolioToolCall(FOLIO_AGENT_TOOL_NAMES.readPage, { page: 1 }, bridge);
    const error = expectError(result);
    expect(error).toContain("read_page");
  });

  test("read_selection and scroll_to_block are also unsupported on the reviewer bridge", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(readFixture());
    const bridge = createReviewerBridge(reviewer);

    expect(
      expectError(executeFolioToolCall(FOLIO_AGENT_TOOL_NAMES.readSelection, {}, bridge)),
    ).toContain("read_selection");
    expect(
      expectError(
        executeFolioToolCall(FOLIO_AGENT_TOOL_NAMES.scrollToBlock, { blockId: "x" }, bridge),
      ),
    ).toContain("scroll_to_block");
  });
});

describe("executeFolioToolCall: happy path against a real FolioDocxReviewer", () => {
  test("list_stories returns typed handles that read_story accepts", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(readFixture());
    const bridge = createReviewerBridge(reviewer);
    const stories = expectOk(
      executeFolioToolCall(FOLIO_AGENT_TOOL_NAMES.listStories, {}, bridge),
    ) as { handle: { type: string }; text: string }[];
    expect(stories.at(0)?.handle).toEqual({ type: "main" });
    const main = expectOk(
      executeFolioToolCall(FOLIO_AGENT_TOOL_NAMES.readStory, { handle: { type: "main" } }, bridge),
    );
    expect(main).toEqual(stories.at(0));
  });

  test("read_document -> find_text -> suggest_changes -> read_changes -> add_comment -> read_comments -> reply_comment -> resolve_comment", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(readFixture(), { author: "AI Reviewer" });
    const bridge = createReviewerBridge(reviewer);

    // read_document
    const blocks = expectOk(
      executeFolioToolCall(FOLIO_AGENT_TOOL_NAMES.readDocument, {}, bridge),
    ) as FolioAgentBlock[];
    expect(blocks.length).toBeGreaterThan(0);
    const heading = blocks.find((block) => block.text.includes("Heading"));
    if (!heading) {
      throw new Error("expected a block containing 'Heading'");
    }

    // find_text
    const findTextResult = expectOk(
      executeFolioToolCall(FOLIO_AGENT_TOOL_NAMES.findText, { query: "heading" }, bridge),
    ) as FolioAgentFindTextResult;
    expect(findTextResult.truncated).toBe(false);
    expect(findTextResult.totalMatches).toBe(findTextResult.matches.length);
    expect(findTextResult.matches.some((match) => match.blockId === heading.blockId)).toBe(true);

    // suggest_changes: replaceInBlock
    const suggestResult = expectOk(
      executeFolioToolCall(
        FOLIO_AGENT_TOOL_NAMES.suggestChanges,
        {
          operations: [
            { type: "replaceInBlock", blockId: heading.blockId, find: "Heading", replace: "Intro" },
          ],
        },
        bridge,
      ),
    ) as FolioAgentApplyOperationsSummary;
    expect(suggestResult.version).toBe(1);
    expect(suggestResult.applied).toHaveLength(1);
    expect(suggestResult.applied[0]?.id).toBe("op-1");
    expect(suggestResult.skipped).toEqual([]);
    expect(suggestResult.receipts).toEqual([
      {
        operationId: "op-1",
        operationIndex: 0,
        affected: [
          {
            type: "block",
            story: "main",
            blockId: heading.blockId,
            effect: "updated",
          },
        ],
      },
    ]);

    // read_changes: the replace is now a pending tracked change
    const changes = expectOk(
      executeFolioToolCall(FOLIO_AGENT_TOOL_NAMES.readChanges, {}, bridge),
    ) as {
      type: string;
      blockId: string | null;
    }[];
    expect(changes.length).toBeGreaterThanOrEqual(2);
    expect(changes.some((change) => change.type === "insertion")).toBe(true);
    expect(changes.some((change) => change.type === "deletion")).toBe(true);

    // add_comment
    const addCommentResult = expectOk(
      executeFolioToolCall(
        FOLIO_AGENT_TOOL_NAMES.addComment,
        { blockId: heading.blockId, text: "Please clarify." },
        bridge,
      ),
    ) as FolioAgentApplyOperationsSummary;
    expect(addCommentResult.version).toBe(1);
    expect(addCommentResult.applied).toHaveLength(1);
    expect(addCommentResult.skipped).toEqual([]);
    expect(addCommentResult.receipts).toEqual([
      {
        operationId: "comment-1",
        operationIndex: 0,
        affected: [
          {
            type: "block",
            story: "main",
            blockId: heading.blockId,
            effect: "commented",
          },
          { type: "comment", commentId: expect.any(Number) },
        ],
      },
    ]);

    // read_comments
    const comments = expectOk(
      executeFolioToolCall(FOLIO_AGENT_TOOL_NAMES.readComments, {}, bridge),
    ) as FolioAgentComment[];
    expect(comments).toHaveLength(1);
    const comment = comments[0];
    if (!comment) {
      throw new Error("expected a comment");
    }
    expect(comment.text).toBe("Please clarify.");
    expect(comment.resolved).toBe(false);

    // reply_comment
    const replyResult = expectOk(
      executeFolioToolCall(
        FOLIO_AGENT_TOOL_NAMES.replyComment,
        { commentId: comment.id, text: "Clarifying now." },
        bridge,
      ),
    ) as { replied: boolean };
    expect(replyResult.replied).toBe(true);

    const afterReply = expectOk(
      executeFolioToolCall(FOLIO_AGENT_TOOL_NAMES.readComments, {}, bridge),
    ) as FolioAgentComment[];
    expect(afterReply[0]?.replies).toHaveLength(1);
    expect(afterReply[0]?.replies[0]?.text).toBe("Clarifying now.");

    // reply_comment against an unknown id fails cleanly
    expect(
      expectError(
        executeFolioToolCall(
          FOLIO_AGENT_TOOL_NAMES.replyComment,
          { commentId: "999999", text: "orphan" },
          bridge,
        ),
      ),
    ).toContain("999999");

    // reply_comment with a malformed id (digits plus trailing junk) is
    // rejected outright rather than truncated to the leading digits and
    // matched against a real comment.
    expect(
      expectError(
        executeFolioToolCall(
          FOLIO_AGENT_TOOL_NAMES.replyComment,
          { commentId: `${comment.id}abc`, text: "orphan" },
          bridge,
        ),
      ),
    ).toContain(`${comment.id}abc`);

    // resolve_comment
    const resolveResult = expectOk(
      executeFolioToolCall(
        FOLIO_AGENT_TOOL_NAMES.resolveComment,
        { commentId: comment.id },
        bridge,
      ),
    ) as { resolved: boolean };
    expect(resolveResult.resolved).toBe(true);

    const afterResolve = expectOk(
      executeFolioToolCall(FOLIO_AGENT_TOOL_NAMES.readComments, { filter: "resolved" }, bridge),
    ) as FolioAgentComment[];
    expect(afterResolve).toHaveLength(1);
    const afterResolveOpen = expectOk(
      executeFolioToolCall(FOLIO_AGENT_TOOL_NAMES.readComments, { filter: "open" }, bridge),
    ) as FolioAgentComment[];
    expect(afterResolveOpen).toHaveLength(0);
  });

  test("suggest_changes reports a plain-language skip reason for a missing block", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(readFixture());
    const bridge = createReviewerBridge(reviewer);

    const result = expectOk(
      executeFolioToolCall(
        FOLIO_AGENT_TOOL_NAMES.suggestChanges,
        {
          operations: [
            {
              id: "custom-1",
              type: "replaceInBlock",
              blockId: "no-such-block",
              find: "x",
              replace: "y",
            },
          ],
        },
        bridge,
      ),
    ) as { applied: unknown[]; receipts: unknown[]; skipped: { id: string; reason: string }[] };
    expect(result.applied).toEqual([]);
    expect(result.receipts).toEqual([]);
    expect(result.skipped).toHaveLength(1);
    expect(result.skipped[0]?.id).toBe("custom-1");
    expect(result.skipped[0]?.reason).toContain("re-read the document");
  });

  test("suggest_changes attaches a block-text precondition and explains a stale target", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(readFixture());
    const reviewerBridge = createReviewerBridge(reviewer);
    const heading = reviewer.snapshot().blocks.find(({ text }) => text.includes("Heading"));
    if (!heading) {
      throw new Error("expected a heading block");
    }
    let receivedPrecondition: { blockTextHash: string } | undefined;
    const bridge = {
      ...reviewerBridge,
      applyDocumentOperations: (
        batch: Parameters<typeof reviewerBridge.applyDocumentOperations>[0],
      ) => {
        receivedPrecondition = batch.operations.at(0)?.precondition;
        return {
          version: batch.version,
          status: "committed" as const,
          applied: [],
          skipped: [{ id: "custom-1", reason: "preconditionFailed" as const }],
        };
      },
    };

    const result = expectOk(
      executeFolioToolCall(
        FOLIO_AGENT_TOOL_NAMES.suggestChanges,
        {
          operations: [
            {
              id: "custom-1",
              type: "replaceInBlock",
              blockId: heading.id,
              find: "Heading",
              replace: "Intro",
            },
          ],
        },
        bridge,
      ),
    ) as { issues: unknown[]; receipts: unknown[]; skipped: { reason: string }[] };

    expect(receivedPrecondition).toEqual({
      blockTextHash: reviewer.snapshot().anchors[heading.id]?.textHash,
    });
    expect(result.skipped[0]?.reason).toContain("re-read the document");
    expect(result.issues).toEqual([]);
    expect(result.receipts).toEqual([]);
  });

  test("suggest_changes explains an unsupported mutation mode", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(readFixture());
    const reviewerBridge = createReviewerBridge(reviewer);
    const bridge = {
      ...reviewerBridge,
      applyDocumentOperations: (
        batch: Parameters<typeof reviewerBridge.applyDocumentOperations>[0],
      ) => ({
        version: batch.version,
        status: "committed" as const,
        applied: [],
        skipped: batch.operations.map(({ id }) => ({ id, reason: "unsupportedMode" as const })),
      }),
    };

    const result = expectOk(
      executeFolioToolCall(
        FOLIO_AGENT_TOOL_NAMES.suggestChanges,
        {
          operations: [
            {
              id: "custom-1",
              type: "replaceInBlock",
              blockId: "seq-0001",
              find: "x",
              replace: "y",
            },
          ],
        },
        bridge,
      ),
    ) as { skipped: { reason: string }[] };

    expect(result.skipped[0]?.reason).toContain("mutation mode");
  });

  test("suggest_changes rejects an empty operations array", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(readFixture());
    const bridge = createReviewerBridge(reviewer);

    const result = executeFolioToolCall(
      FOLIO_AGENT_TOOL_NAMES.suggestChanges,
      { operations: [] },
      bridge,
    );
    expect(expectError(result)).toContain("empty");
  });

  test("suggest_changes rejects an operations array over the per-call cap", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(readFixture());
    const bridge = createReviewerBridge(reviewer);

    const operations = Array.from({ length: 51 }, (_, i) => ({
      type: "replaceInBlock",
      blockId: `block-${i}`,
      find: "x",
      replace: "y",
    }));
    const result = executeFolioToolCall(
      FOLIO_AGENT_TOOL_NAMES.suggestChanges,
      { operations },
      bridge,
    );
    expect(expectError(result)).toContain("50-operation limit");
  });

  test("suggest_changes rejects find/replace/text/comment strings over the length cap", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(readFixture());
    const bridge = createReviewerBridge(reviewer);
    const tooLong = "x".repeat(100_001);

    const findTooLong = executeFolioToolCall(
      FOLIO_AGENT_TOOL_NAMES.suggestChanges,
      { operations: [{ type: "replaceInBlock", blockId: "b1", find: tooLong, replace: "y" }] },
      bridge,
    );
    expect(expectError(findTooLong)).toContain("100,000-character limit");

    const replaceTooLong = executeFolioToolCall(
      FOLIO_AGENT_TOOL_NAMES.suggestChanges,
      { operations: [{ type: "replaceInBlock", blockId: "b1", find: "x", replace: tooLong }] },
      bridge,
    );
    expect(expectError(replaceTooLong)).toContain("100,000-character limit");

    const textTooLong = executeFolioToolCall(
      FOLIO_AGENT_TOOL_NAMES.suggestChanges,
      { operations: [{ type: "replaceBlock", blockId: "b1", text: tooLong }] },
      bridge,
    );
    expect(expectError(textTooLong)).toContain("100,000-character limit");

    const commentTooLong = executeFolioToolCall(
      FOLIO_AGENT_TOOL_NAMES.suggestChanges,
      { operations: [{ type: "deleteBlock", blockId: "b1", comment: tooLong }] },
      bridge,
    );
    expect(expectError(commentTooLong)).toContain("100,000-character limit");
  });

  test("add_comment rejects text/quote over the length cap", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(readFixture());
    const bridge = createReviewerBridge(reviewer);
    const tooLong = "x".repeat(100_001);
    const blocks = reviewer.snapshot().blocks;
    const target = blocks[0];
    if (!target) {
      throw new Error("expected at least one block");
    }

    expect(
      expectError(
        executeFolioToolCall(
          FOLIO_AGENT_TOOL_NAMES.addComment,
          { blockId: target.id, text: tooLong },
          bridge,
        ),
      ),
    ).toContain("100,000-character limit");

    expect(
      expectError(
        executeFolioToolCall(
          FOLIO_AGENT_TOOL_NAMES.addComment,
          { blockId: target.id, text: "fine", quote: tooLong },
          bridge,
        ),
      ),
    ).toContain("100,000-character limit");
  });

  test("reply_comment rejects text over the length cap", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(readFixture());
    const bridge = createReviewerBridge(reviewer);
    const tooLong = "x".repeat(100_001);

    const result = executeFolioToolCall(
      FOLIO_AGENT_TOOL_NAMES.replyComment,
      { commentId: "1", text: tooLong },
      bridge,
    );
    expect(expectError(result)).toContain("100,000-character limit");
  });
});

describe("find_text edge cases", () => {
  test("multiple occurrences in one block increment occurrenceInBlock, matching is case-insensitive by default", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(readFixture());
    const bridge = createReviewerBridge(reviewer);

    // Apply an edit that guarantees a repeated substring inside a single block.
    const blocks = reviewer.snapshot().blocks;
    const target = blocks.find((block) => block.text.includes("Heading"));
    if (!target) {
      throw new Error("expected a block containing 'Heading'");
    }
    reviewer.applyOperations(
      [{ id: "r1", type: "replaceBlock", blockId: target.id, text: "repeat repeat REPEAT" }],
      { mode: "direct" },
    );

    const findTextResult = expectOk(
      executeFolioToolCall(FOLIO_AGENT_TOOL_NAMES.findText, { query: "repeat" }, bridge),
    ) as FolioAgentFindTextResult;
    expect(findTextResult.matches).toHaveLength(3);
    expect(findTextResult.truncated).toBe(false);
    expect(findTextResult.totalMatches).toBe(3);
    expect(findTextResult.matches.map((match) => match.occurrenceInBlock)).toEqual([0, 1, 2]);
    expect(
      findTextResult.matches.map((match) => [match.range.startOffset, match.range.endOffset]),
    ).toEqual([
      [0, 6],
      [7, 13],
      [14, 20],
    ]);

    const secondMatch = findTextResult.matches.at(1);
    if (!secondMatch) {
      throw new Error("expected the second match");
    }
    const suggested = expectOk(
      executeFolioToolCall(
        FOLIO_AGENT_TOOL_NAMES.suggestChanges,
        { operations: [{ type: "replaceRange", range: secondMatch.range, replace: "done" }] },
        bridge,
      ),
    ) as FolioAgentApplyOperationsSummary;
    expect(suggested.applied).toEqual([{ id: "op-1" }]);

    const caseSensitive = expectOk(
      executeFolioToolCall(
        FOLIO_AGENT_TOOL_NAMES.findText,
        { query: "repeat", matchCase: true },
        bridge,
      ),
    ) as FolioAgentFindTextResult;
    expect(caseSensitive.matches).toHaveLength(1);
    expect(caseSensitive.totalMatches).toBe(1);
  });

  test("query over the length cap is rejected", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(readFixture());
    const bridge = createReviewerBridge(reviewer);

    const result = executeFolioToolCall(
      FOLIO_AGENT_TOOL_NAMES.findText,
      { query: "a".repeat(1_001) },
      bridge,
    );
    expect(expectError(result)).toContain("1,000-character limit");
  });

  test("matches beyond the cap are reported as truncated with an accurate totalMatches", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(readFixture());
    const bridge = createReviewerBridge(reviewer);

    const blocks = reviewer.snapshot().blocks;
    const target = blocks[0];
    if (!target) {
      throw new Error("expected at least one block");
    }
    // 250 occurrences of "x" in one block, over the 200-match cap.
    reviewer.applyOperations(
      [{ id: "r1", type: "replaceBlock", blockId: target.id, text: "x ".repeat(250) }],
      { mode: "direct" },
    );

    const result = expectOk(
      executeFolioToolCall(FOLIO_AGENT_TOOL_NAMES.findText, { query: "x" }, bridge),
    ) as FolioAgentFindTextResult;
    expect(result.matches).toHaveLength(200);
    expect(result.truncated).toBe(true);
    expect(result.totalMatches).toBe(250);
  });

  test("whole-word matching uses Unicode word boundaries", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(readFixture());
    const bridge = createReviewerBridge(reviewer);
    const target = reviewer.snapshot().blocks.at(0);
    if (target === undefined) {
      throw new Error("expected at least one block");
    }
    reviewer.applyOperations(
      [
        {
          id: "unicode",
          type: "replaceBlock",
          blockId: target.id,
          text: "žaloba předžaloba ŽALOBA",
        },
      ],
      { mode: "direct" },
    );

    const result = expectOk(
      executeFolioToolCall(
        FOLIO_AGENT_TOOL_NAMES.findText,
        { query: "žaloba", wholeWord: true },
        bridge,
      ),
    ) as FolioAgentFindTextResult;

    expect(result.totalMatches).toBe(2);
  });

  test("whole-word matching stays correct for a match far into a very large block", async () => {
    // Regression guard for bounding find_text's whole-word boundary check to
    // a small fixed window instead of slicing the whole block on either side
    // of every match (previously O(block.text.length) per match). Both
    // occurrences sit tens of thousands of characters into the block;
    // "prefixedTARGET" directly abuts a word character on its left, so the
    // exclusion must still trigger from a window that only looks a few
    // characters either side of the match.
    const reviewer = await FolioDocxReviewer.fromBuffer(readFixture());
    const bridge = createReviewerBridge(reviewer);
    const target = reviewer.snapshot().blocks.at(0);
    if (target === undefined) {
      throw new Error("expected at least one block");
    }
    const filler = "x".repeat(20_000);
    const text = `${filler} TARGET ${filler} prefixedTARGET ${filler}`;
    reviewer.applyOperations(
      [{ id: "large-block", type: "replaceBlock", blockId: target.id, text }],
      { mode: "direct" },
    );

    const result = expectOk(
      executeFolioToolCall(
        FOLIO_AGENT_TOOL_NAMES.findText,
        { query: "TARGET", wholeWord: true },
        bridge,
      ),
    ) as FolioAgentFindTextResult;

    expect(result.totalMatches).toBe(1);
  });

  test("limits main-story search to real page mappings and exposes the page on matches", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(readFixture());
    const base = createReviewerBridge(reviewer);
    const target = reviewer.snapshot().blocks.find(({ text }) => text.includes("Heading"));
    if (target === undefined) {
      throw new Error("expected a heading block");
    }
    const bridge: FolioAgentBridge = {
      ...base,
      getPageCount: () => 2,
      getTargetPage: (received) =>
        received.type === "textRange" && received.blockId === target.id ? 2 : 1,
    };

    const result = expectOk(
      executeFolioToolCall(
        FOLIO_AGENT_TOOL_NAMES.findText,
        { query: "Heading", scope: { type: "page", page: 2 } },
        bridge,
      ),
    ) as FolioAgentFindTextResult;

    expect(result.totalMatches).toBeGreaterThan(0);
    expect(result.matches.every((match) => match.type === "main" && match.page === 2)).toBe(true);
  });

  test("searches a non-main story without inventing block handles", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(readFixture());
    const base = createReviewerBridge(reviewer);
    const bridge: FolioAgentBridge = {
      ...base,
      readStory: (handle) =>
        handle.type === "header"
          ? { handle, text: "Privileged header text" }
          : (base.readStory?.(handle) ?? null),
    };

    const result = expectOk(
      executeFolioToolCall(
        FOLIO_AGENT_TOOL_NAMES.findText,
        {
          query: "header",
          scope: { type: "story", handle: { type: "header", relationshipId: "rId7" } },
        },
        bridge,
      ),
    ) as FolioAgentStoryFindTextResult;

    expect(result.matches.at(0)).toMatchObject({
      type: "story",
      startOffset: 11,
      endOffset: 17,
    });
  });
});

describe("outline-first scoped reads", () => {
  test("returns stable outline handles and bounded section pages", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(readFixture());
    const base = createReviewerBridge(reviewer);
    const snapshot = reviewer.snapshot();
    const firstBlock = snapshot.blocks.at(0);
    if (firstBlock === undefined) {
      throw new Error("expected fixture to contain a block");
    }
    const bridge: FolioAgentBridge = {
      ...base,
      snapshot: () => ({
        ...snapshot,
        blocks: [{ ...firstBlock, kind: "heading", headingLevel: 1 }, ...snapshot.blocks.slice(1)],
      }),
    };
    const outline = expectOk(
      executeFolioToolCall(FOLIO_AGENT_TOOL_NAMES.getDocumentOutline, {}, bridge),
    ) as FolioAgentDocumentOutline;
    const first = outline.sections.at(0);
    if (first === undefined) {
      throw new Error("expected fixture to contain an outline heading");
    }

    const firstPage = expectOk(
      executeFolioToolCall(
        FOLIO_AGENT_TOOL_NAMES.readSection,
        { handle: first.handle, maxBlocks: 1 },
        bridge,
      ),
    ) as FolioAgentSectionRead;
    expect(firstPage.blocks).toHaveLength(1);
    expect(firstPage.totalBlocks).toBeGreaterThanOrEqual(1);

    if (firstPage.nextAfterBlockId !== undefined) {
      const secondPage = expectOk(
        executeFolioToolCall(
          FOLIO_AGENT_TOOL_NAMES.readSection,
          { handle: first.handle, maxBlocks: 1, afterBlockId: firstPage.nextAfterBlockId },
          bridge,
        ),
      ) as FolioAgentSectionRead;
      expect(secondPage.blocks.at(0)?.blockId).not.toBe(firstPage.blocks.at(0)?.blockId);
    }
  });

  test("show_in_document accepts either a block or an exact range", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(readFixture());
    const base = createReviewerBridge(reviewer);
    const block = reviewer.snapshot().blocks.at(0);
    if (block === undefined) {
      throw new Error("expected at least one block");
    }
    const targets: unknown[] = [];
    const bridge: FolioAgentBridge = {
      ...base,
      showInDocument: (target) => {
        targets.push(target);
        return true;
      },
    };
    const match = (
      expectOk(
        executeFolioToolCall(
          FOLIO_AGENT_TOOL_NAMES.findText,
          { query: block.text.slice(0, 2) },
          bridge,
        ),
      ) as FolioAgentFindTextResult
    ).matches.at(0);
    if (match === undefined || match.type !== "main") {
      throw new Error("expected a main-story match");
    }

    expectOk(
      executeFolioToolCall(FOLIO_AGENT_TOOL_NAMES.showInDocument, { blockId: block.id }, bridge),
    );
    expectOk(
      executeFolioToolCall(FOLIO_AGENT_TOOL_NAMES.showInDocument, { range: match.range }, bridge),
    );
    expect(targets).toEqual([{ type: "block", story: "main", blockId: block.id }, match.range]);
  });
});
