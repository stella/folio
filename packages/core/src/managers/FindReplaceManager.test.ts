import { describe, expect, test } from "bun:test";

import { FindReplaceManager, getAdjacentFindIndex } from "./FindReplaceManager";
import type { OpStory } from "@stll/docx-core/ops";
import type { Document } from "../types/document";

type TestMatch = {
  paragraphIndex: number;
  startOffset: number;
  endOffset: number;
  id: string;
};

const match = (id: string, paragraphIndex: number, startOffset: number): TestMatch => ({
  id,
  paragraphIndex,
  startOffset,
  endOffset: startOffset + 1,
});

type StoryTestMatch = TestMatch & {
  story: OpStory;
  textBoxIndex?: number;
  commentId?: number;
};

const unsupportedMatches = [
  { ...match("header", 0, 0), story: { kind: "header", rId: "rId1" } },
  { ...match("footer", 0, 0), story: { kind: "footer", rId: "rId2" } },
  { ...match("footnote", 0, 0), story: { kind: "footnote", id: 1 } },
  { ...match("endnote", 0, 0), story: { kind: "endnote", id: 2 } },
  { ...match("shape", 0, 0), story: "main", textBoxIndex: 0 },
  { ...match("comment", 0, 0), story: "main", commentId: 7 },
] as const satisfies readonly StoryTestMatch[];

describe("FindReplaceManager replacement story boundaries", () => {
  test.each(unsupportedMatches)(
    "keeps $id navigable while refusing body replacement",
    (secondary) => {
      const document: Document = {
        package: {
          document: {
            content: [
              {
                type: "paragraph",
                content: [{ type: "run", content: [{ type: "text", text: "body" }] }],
              },
            ],
          },
        },
      };
      const before = structuredClone(document);
      const body = { ...match("body", 0, 0), story: "main" } as const satisfies StoryTestMatch;
      const manager = new FindReplaceManager<StoryTestMatch>();
      manager.setMatches([body, secondary]);
      expect(manager.goTo(1)?.match).toEqual(secondary);
      expect(manager.replaceCurrent(document, "changed")).toBeNull();
      expect(manager.replaceAll(document, [body, secondary], "changed")).toBeNull();
      expect(document).toEqual(before);
      expect(manager.navigate("previous")?.match).toEqual(body);
      expect(manager.replaceCurrent(document, "B")).not.toBeNull();
    },
  );
});

describe("getAdjacentFindIndex", () => {
  test("wraps forward past the last match to the first", () => {
    expect(getAdjacentFindIndex(2, 3, "next")).toBe(0);
  });

  test("wraps backward past the first match to the last", () => {
    expect(getAdjacentFindIndex(0, 3, "previous")).toBe(2);
  });

  test("steps within range", () => {
    expect(getAdjacentFindIndex(0, 3, "next")).toBe(1);
    expect(getAdjacentFindIndex(2, 3, "previous")).toBe(1);
  });

  test("returns 0 when there are no matches", () => {
    expect(getAdjacentFindIndex(0, 0, "next")).toBe(0);
    expect(getAdjacentFindIndex(0, 0, "previous")).toBe(0);
  });
});

describe("FindReplaceManager cursor", () => {
  test("setMatches seeds the cursor on the first match and keeps an empty result", () => {
    const manager = new FindReplaceManager<TestMatch>();
    const matches = [match("a", 0, 0), match("b", 0, 4)];

    expect(manager.setMatches(matches)).toEqual({ matches, totalCount: 2, currentIndex: 0 });

    expect(manager.setMatches([])).toEqual({ matches: [], totalCount: 0, currentIndex: 0 });
  });

  test("navigate wraps the cursor and reports the now-current match", () => {
    const manager = new FindReplaceManager<TestMatch>();
    const matches = [match("a", 0, 0), match("b", 1, 0)];
    manager.setMatches(matches);

    expect(manager.navigate("next")).toEqual({ match: matches[1]!, index: 1 });
    expect(manager.navigate("next")).toEqual({ match: matches[0]!, index: 0 });
    expect(manager.navigate("previous")).toEqual({ match: matches[1]!, index: 1 });
    expect(manager.getResult()?.currentIndex).toBe(1);
  });

  test("navigate returns null with no matches and clear discards the result", () => {
    const manager = new FindReplaceManager<TestMatch>();
    expect(manager.navigate("next")).toBeNull();

    manager.setMatches([match("a", 0, 0)]);
    manager.clear();
    expect(manager.getResult()).toBeNull();
    expect(manager.navigate("next")).toBeNull();
  });

  test("keeps explicit jumps synchronized with subsequent navigation", () => {
    const manager = new FindReplaceManager<TestMatch>();
    const matches = [match("a", 0, 0), match("b", 0, 4), match("c", 0, 8)];
    manager.setMatches(matches);

    expect(manager.goTo(1)).toEqual({ match: matches[1]!, index: 1 });
    expect(manager.navigate("next")).toEqual({ match: matches[2]!, index: 2 });
    expect(manager.goTo(9)).toBeNull();
    expect(manager.getResult()?.currentIndex).toBe(2);
  });
});

// Refreshes keep the current cursor bounded as the result set shrinks.
test("result refresh preserves or clamps the cursor", () => {
  const manager = new FindReplaceManager<TestMatch>();
  const matches = [match("a", 0, 0), match("b", 1, 0), match("c", 2, 0)];
  for (const index of [-1, 0, 1, 2, 3, 20]) {
    const result = manager.setMatches(matches, index);
    expect(result.currentIndex).toBe(Math.max(0, Math.min(index, matches.length - 1)));
    expect(manager.navigate("next")?.index).toBe((result.currentIndex + 1) % matches.length);
  }
  expect(manager.setMatches([], 4).currentIndex).toBe(0);
  expect(manager.navigate("next")).toBeNull();
});
