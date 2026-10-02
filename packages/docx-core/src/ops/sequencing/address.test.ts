import { expect, test } from "bun:test";
import { OP_STORIES } from "../types";
import { isOpStory, sameStory } from "./address";

test("story decoding derives its accepted identities from the operation contract", () => {
  for (const story of Object.values(OP_STORIES)) {
    expect(isOpStory(story)).toBe(true);
    expect(sameStory(story, story)).toBe(true);
  }
  for (const invalid of [undefined, null, false, 0, [], {}, { type: "main" }, "unknown-story"]) {
    expect(isOpStory(invalid)).toBe(false);
  }
});
