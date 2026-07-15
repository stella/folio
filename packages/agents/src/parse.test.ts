/**
 * `parseSuggestChangesInput` / `parseAddCommentInput` unit tests: the
 * validation rules `execute.ts`'s `suggest_changes` / `add_comment` tools run
 * before applying anything, exercised directly here (no bridge, no
 * document) so a host with its own review-queue UX can trust these are the
 * canonical rules. Expectations mirror `execute.test.ts`'s
 * `suggest_changes` / `add_comment` cases, since `execute.ts` now delegates
 * to these same functions.
 */

import { describe, expect, test } from "bun:test";

import { parseAddCommentInput, parseSuggestChangesInput } from "./parse";
import { SUGGEST_CHANGES_OPERATION_TYPES } from "./tools";

describe("parseAddCommentInput", () => {
  test("valid input builds a commentOnBlock operation", () => {
    const result = parseAddCommentInput({ blockId: "b1", text: "Please clarify." });
    expect(result.ok).toBe(true);
    if (!result.ok) {
      throw new Error("expected ok:true");
    }
    expect(result.operation).toEqual({
      id: "comment-1",
      type: "commentOnBlock",
      blockId: "b1",
      comment: { text: "Please clarify." },
    });
  });

  test("valid input with a quote includes it on the operation", () => {
    const result = parseAddCommentInput({ blockId: "b1", text: "note", quote: "the quoted text" });
    expect(result.ok).toBe(true);
    if (!result.ok) {
      throw new Error("expected ok:true");
    }
    expect(result.operation).toMatchObject({ quote: "the quoted text" });
  });

  test("rejects non-object args", () => {
    const result = parseAddCommentInput("not an object");
    expect(result.ok).toBe(false);
  });

  test("rejects a missing blockId", () => {
    const result = parseAddCommentInput({ text: "note" });
    expect(result.ok).toBe(false);
    if (result.ok) {
      throw new Error("expected ok:false");
    }
    expect(result.error).toContain("blockId");
  });

  test("rejects a missing text", () => {
    const result = parseAddCommentInput({ blockId: "b1" });
    expect(result.ok).toBe(false);
    if (result.ok) {
      throw new Error("expected ok:false");
    }
    expect(result.error).toContain("text");
  });

  test("rejects a non-string quote", () => {
    const result = parseAddCommentInput({ blockId: "b1", text: "note", quote: 42 });
    expect(result.ok).toBe(false);
    if (result.ok) {
      throw new Error("expected ok:false");
    }
    expect(result.error).toContain("quote");
  });

  test("rejects text over the length cap", () => {
    const tooLong = "x".repeat(100_001);
    const result = parseAddCommentInput({ blockId: "b1", text: tooLong });
    expect(result.ok).toBe(false);
    if (result.ok) {
      throw new Error("expected ok:false");
    }
    expect(result.error).toContain("100,000-character limit");
  });

  test("rejects quote over the length cap", () => {
    const tooLong = "x".repeat(100_001);
    const result = parseAddCommentInput({ blockId: "b1", text: "fine", quote: tooLong });
    expect(result.ok).toBe(false);
    if (result.ok) {
      throw new Error("expected ok:false");
    }
    expect(result.error).toContain("100,000-character limit");
  });
});

describe("parseSuggestChangesInput", () => {
  test("valid commentOnRange operation preserves the selected range", () => {
    const range = {
      type: "textRange",
      story: "main",
      blockId: "b1",
      startOffset: 0,
      endOffset: 4,
      selectedTextHash: "h123",
    };
    const result = parseSuggestChangesInput({
      operations: [{ type: "commentOnRange", range, comment: "Review this" }],
    });
    expect(result.ok).toBe(true);
    if (!result.ok) {
      throw new Error("expected ok:true");
    }
    expect(result.operations).toEqual([
      { id: "op-1", type: "commentOnRange", range, comment: { text: "Review this" } },
    ]);
  });

  test("valid replaceRange operation preserves the handle returned by find_text", () => {
    const range = {
      type: "textRange",
      story: "main",
      blockId: "b1",
      startOffset: 7,
      endOffset: 13,
      selectedTextHash: "h123",
    };
    const result = parseSuggestChangesInput({
      operations: [{ type: "replaceRange", range, replace: "done" }],
    });
    expect(result.ok).toBe(true);
    if (!result.ok) {
      throw new Error("expected ok:true");
    }
    expect(result.operations).toEqual([
      { id: "op-1", type: "replaceRange", range, replace: "done" },
    ]);
  });

  test("valid replaceInBlock operation is parsed with an auto-generated id", () => {
    const result = parseSuggestChangesInput({
      operations: [{ type: "replaceInBlock", blockId: "b1", find: "Heading", replace: "Intro" }],
    });
    expect(result.ok).toBe(true);
    if (!result.ok) {
      throw new Error("expected ok:true");
    }
    expect(result.operations).toEqual([
      { id: "op-1", type: "replaceInBlock", blockId: "b1", find: "Heading", replace: "Intro" },
    ]);
  });

  test("a caller-supplied id is preserved instead of the auto-generated one", () => {
    const result = parseSuggestChangesInput({
      operations: [{ id: "custom-1", type: "deleteBlock", blockId: "b1" }],
    });
    expect(result.ok).toBe(true);
    if (!result.ok) {
      throw new Error("expected ok:true");
    }
    expect(result.operations[0]?.id).toBe("custom-1");
  });

  test("an optional comment attaches a review comment to the operation", () => {
    const result = parseSuggestChangesInput({
      operations: [{ type: "deleteBlock", blockId: "b1", comment: "why this edit" }],
    });
    expect(result.ok).toBe(true);
    if (!result.ok) {
      throw new Error("expected ok:true");
    }
    expect(result.operations[0]).toMatchObject({ comment: { text: "why this edit" } });
  });

  test("rejects a missing operations array", () => {
    const result = parseSuggestChangesInput({});
    expect(result.ok).toBe(false);
    if (result.ok) {
      throw new Error("expected ok:false");
    }
    expect(result.error).toContain("operations");
  });

  test("rejects an empty operations array", () => {
    const result = parseSuggestChangesInput({ operations: [] });
    expect(result.ok).toBe(false);
    if (result.ok) {
      throw new Error("expected ok:false");
    }
    expect(result.error).toContain("empty");
  });

  test("rejects an operations array over the per-call cap", () => {
    const operations = Array.from({ length: 51 }, (_, i) => ({
      type: "replaceInBlock",
      blockId: `block-${i}`,
      find: "x",
      replace: "y",
    }));
    const result = parseSuggestChangesInput({ operations });
    expect(result.ok).toBe(false);
    if (result.ok) {
      throw new Error("expected ok:false");
    }
    expect(result.error).toContain("50-operation limit");
  });

  test("rejects an unknown operation type", () => {
    const result = parseSuggestChangesInput({
      operations: [{ type: "notARealType", blockId: "b1" }],
    });
    expect(result.ok).toBe(false);
    if (result.ok) {
      throw new Error("expected ok:false");
    }
    expect(result.error).toContain("type");
  });

  test("rejects a replaceInBlock operation missing find/replace", () => {
    const result = parseSuggestChangesInput({
      operations: [{ type: "replaceInBlock", blockId: "b1" }],
    });
    expect(result.ok).toBe(false);
    if (result.ok) {
      throw new Error("expected ok:false");
    }
    expect(result.error).toContain("find");
  });

  test("rejects find/replace/text/comment strings over the length cap", () => {
    const tooLong = "x".repeat(100_001);

    const findTooLong = parseSuggestChangesInput({
      operations: [{ type: "replaceInBlock", blockId: "b1", find: tooLong, replace: "y" }],
    });
    expect(findTooLong.ok).toBe(false);
    if (findTooLong.ok) {
      throw new Error("expected ok:false");
    }
    expect(findTooLong.error).toContain("100,000-character limit");

    const replaceTooLong = parseSuggestChangesInput({
      operations: [{ type: "replaceInBlock", blockId: "b1", find: "x", replace: tooLong }],
    });
    expect(replaceTooLong.ok).toBe(false);
    if (replaceTooLong.ok) {
      throw new Error("expected ok:false");
    }
    expect(replaceTooLong.error).toContain("100,000-character limit");

    const textTooLong = parseSuggestChangesInput({
      operations: [{ type: "replaceBlock", blockId: "b1", text: tooLong }],
    });
    expect(textTooLong.ok).toBe(false);
    if (textTooLong.ok) {
      throw new Error("expected ok:false");
    }
    expect(textTooLong.error).toContain("100,000-character limit");

    const commentTooLong = parseSuggestChangesInput({
      operations: [{ type: "deleteBlock", blockId: "b1", comment: tooLong }],
    });
    expect(commentTooLong.ok).toBe(false);
    if (commentTooLong.ok) {
      throw new Error("expected ok:false");
    }
    expect(commentTooLong.error).toContain("100,000-character limit");
  });

  test("accepts every operation type the suggest_changes tool schema advertises", () => {
    // Guards the tool-schema derivation in tools.ts: since the schema's type
    // enum is computed from the contract's operation-type list minus an
    // exclusion list, a new contract type would silently appear in the enum
    // even though this parser rejects it. A superset fixture works for every
    // type because parseSuggestChangesInput ignores fields a type does not use.
    const supersetOperation = (type: string) => ({
      type,
      blockId: "b1",
      find: "old",
      replace: "new",
      text: "inserted",
      comment: "why",
      range: {
        type: "textRange",
        story: "main",
        blockId: "b1",
        startOffset: 0,
        endOffset: 4,
        selectedTextHash: "h123",
      },
    });
    for (const type of SUGGEST_CHANGES_OPERATION_TYPES) {
      const result = parseSuggestChangesInput({ operations: [supersetOperation(type)] });
      expect(result.ok, type).toBe(true);
    }
    for (const excludedType of [
      "formatRange",
      "commentOnBlock",
      "insertSignatureTable",
      "insertTableRow",
      "deleteTableRow",
    ]) {
      const result = parseSuggestChangesInput({ operations: [supersetOperation(excludedType)] });
      expect(result.ok, excludedType).toBe(false);
    }
  });
});
