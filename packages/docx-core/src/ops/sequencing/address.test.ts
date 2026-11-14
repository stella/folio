import { expect, test } from "bun:test";

import { OP_STORIES } from "../types";
import { isOpStory, sameStory } from "./address";

test("story decoding accepts every story the operation contract names", () => {
  for (const story of [
    ...Object.values(OP_STORIES),
    { kind: "header", rId: "rId7" },
    { kind: "footer", rId: "rId8" },
    { kind: "footnote", id: 1 },
    { kind: "endnote", id: 0 },
  ] as const) {
    expect(isOpStory(story)).toBe(true);
    expect(sameStory(story, story)).toBe(true);
  }
  for (const invalid of [
    undefined,
    null,
    false,
    0,
    [],
    {},
    { type: "main" },
    "unknown-story",
    { kind: "header" },
    { kind: "header", rId: 7 },
    { kind: "header", rId: "rId7", id: 1 },
    { kind: "body", rId: "rId7" },
    { kind: "footnote", id: -1 },
    { kind: "footnote", id: 1.5 },
    { kind: "footnote", id: "1" },
    { kind: "endnote", id: 1, rId: "rId7" },
    { kind: "comment", id: 1 },
  ]) {
    expect(isOpStory(invalid)).toBe(false);
  }
});
