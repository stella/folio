import { describe, expect, test } from "bun:test";

import {
  decodeCommentId,
  decodeMainStoryTextRangeHandle,
  decodeSectionHandle,
  decodeStoryHandle,
} from "./codecs";

describe("boundary codecs", () => {
  test("decodeCommentId accepts bare digits and rejects malformed strings", () => {
    expect(decodeCommentId("17")).toEqual({ raw: "17", numeric: 17 });
    expect(decodeCommentId("17abc")).toBeNull();
    expect(decodeCommentId("-17")).toBeNull();
    expect(decodeCommentId(17)).toBeNull();
    expect(decodeCommentId("9007199254740992")).toBeNull();
    expect(decodeCommentId("9".repeat(400))).toBeNull();
  });

  test("decodeStoryHandle only accepts valid typed story handles", () => {
    expect(decodeStoryHandle({ type: "main" })).toEqual({ type: "main" });
    expect(decodeStoryHandle({ type: "header", relationshipId: "rId7" })).toEqual({
      type: "header",
      relationshipId: "rId7",
    });
    expect(decodeStoryHandle({ type: "endnote", noteId: 4 })).toEqual({
      type: "endnote",
      noteId: 4,
    });
    expect(decodeStoryHandle({ type: "header" })).toBeNull();
    expect(decodeStoryHandle({ type: "footnote", noteId: 1.5 })).toBeNull();
    expect(decodeStoryHandle({ type: "unknown" })).toBeNull();
  });

  test("decodeSectionHandle only accepts main-story heading-section handles", () => {
    expect(
      decodeSectionHandle({
        type: "headingSection",
        story: "main",
        headingBlockId: "block-3",
        headingTextHash: "habc123",
        headingLevel: 2,
      }),
    ).toEqual({
      type: "headingSection",
      story: "main",
      headingBlockId: "block-3",
      headingTextHash: "habc123",
      headingLevel: 2,
    });
    expect(
      decodeSectionHandle({
        type: "headingSection",
        story: "header",
        headingBlockId: "block-3",
        headingTextHash: "habc123",
        headingLevel: 2,
      }),
    ).toBeNull();
    expect(
      decodeSectionHandle({
        type: "headingSection",
        story: "main",
        headingBlockId: "block-3",
        headingTextHash: "habc123",
        headingLevel: 0,
      }),
    ).toBeNull();
    expect(
      decodeSectionHandle({
        type: "headingSection",
        story: "main",
        headingBlockId: "block-3",
        headingTextHash: "not-a-hash",
        headingLevel: 2,
      }),
    ).toBeNull();
  });

  test("decodeMainStoryTextRangeHandle rejects hashes and offsets that cannot round-trip from find_text", () => {
    expect(
      decodeMainStoryTextRangeHandle({
        type: "textRange",
        story: "main",
        blockId: "block-9",
        startOffset: 3,
        endOffset: 8,
        selectedTextHash: "h1z9",
      }),
    ).toEqual({
      type: "textRange",
      story: "main",
      blockId: "block-9",
      startOffset: 3,
      endOffset: 8,
      selectedTextHash: "h1z9",
    });
    expect(
      decodeMainStoryTextRangeHandle({
        type: "textRange",
        story: "main",
        blockId: "block-9",
        startOffset: 8,
        endOffset: 8,
        selectedTextHash: "h1z9",
      }),
    ).toBeNull();
    expect(
      decodeMainStoryTextRangeHandle({
        type: "textRange",
        story: "main",
        blockId: "block-9",
        startOffset: 3,
        endOffset: 8,
        selectedTextHash: "not-a-hash",
      }),
    ).toBeNull();
  });
});
