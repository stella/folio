import { expect, test } from "bun:test";

import type { ParagraphBlock } from "../types";
import {
  clearTextWidthCache,
  getCachedTextWidth,
  getTextCacheSize,
  hashParagraphBlock,
  setCachedTextWidth,
  setTextCacheSize,
} from "./cache";

test("text width cache retains frequently reused entries when it evicts", () => {
  try {
    setTextCacheSize(2);
    clearTextWidthCache();
    setCachedTextWidth("hot", "16px Arial", 0, 1);
    setCachedTextWidth("cold", "16px Arial", 0, 2);

    for (let hit = 0; hit < 32; hit += 1) {
      expect(getCachedTextWidth("hot", "16px Arial", 0)).toBe(1);
    }
    setCachedTextWidth("new", "16px Arial", 0, 3);

    expect(getCachedTextWidth("hot", "16px Arial", 0)).toBe(1);
    expect(getCachedTextWidth("cold", "16px Arial", 0)).toBeUndefined();
    expect(getCachedTextWidth("new", "16px Arial", 0)).toBe(3);
    expect(getTextCacheSize()).toBe(2);
  } finally {
    setTextCacheSize(20_000);
    clearTextWidthCache();
  }
});

test("paragraph cache keys include complex-script list marker formatting", () => {
  const paragraphWithMarkerSize = (complexScriptFontSize: number): ParagraphBlock => ({
    kind: "paragraph",
    id: "list-item",
    runs: [{ kind: "text", text: "بند" }],
    attrs: {
      listMarker: "ا.",
      listMarkerFormatting: {
        complexScriptFontFamily: "Traditional Arabic",
        complexScriptFontSize,
      },
    },
  });

  expect(hashParagraphBlock(paragraphWithMarkerSize(12))).not.toBe(
    hashParagraphBlock(paragraphWithMarkerSize(18)),
  );
});
