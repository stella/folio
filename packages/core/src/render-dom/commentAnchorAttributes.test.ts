import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { afterAll, describe, expect, test } from "bun:test";

import {
  commentAnchorIds,
  commentAnchorSelector,
  indexCommentAnchors,
  writeCommentAnchorIds,
} from "./commentAnchorAttributes";

afterAll(() => {
  GlobalRegistrator.unregister();
});

/** Two runs: the first only in comment 7, the second in both 7 and 9. */
const paintedRuns = (): HTMLElement => {
  const root = document.createElement("div");
  for (const commentIds of [[7], [7, 9], [9]]) {
    const run = document.createElement("span");
    writeCommentAnchorIds(run, commentIds);
    root.append(run);
  }
  return root;
};

describe("comment anchor attributes", () => {
  test("a selector for either comment finds the run inside both ranges", () => {
    const root = paintedRuns();
    const [, overlapping] = [...root.children];

    expect([...root.querySelectorAll(commentAnchorSelector(7))]).toContain(overlapping);
    expect([...root.querySelectorAll(commentAnchorSelector(9))]).toContain(overlapping);
  });

  test("the index answers for every comment a run is in, first run first", () => {
    const root = paintedRuns();
    const [first, overlapping] = [...root.children];

    const anchors = indexCommentAnchors(root);
    expect(anchors.get("7")).toBe(first);
    // Comment 9's range opens at the overlap, not at the run whose own id it is.
    expect(anchors.get("9")).toBe(overlapping);
  });

  test("a host-supplied id cannot break out of the selector", () => {
    // `comment.id` is typed as a number, but a controlled `comments` prop is
    // not runtime-checked. Nothing may close the attribute value early.
    const selector = commentAnchorSelector('1"] , img[src=x onerror=alert(1)]');

    // The payload survives only escaped: it never closes the attribute value.
    expect(selector).not.toContain('1"] , img[');
    expect(selector).toContain("data-comment-ids~=");
  });

  test("reading an anchor with no comment yields no ids", () => {
    const empty = document.createElement("span");
    writeCommentAnchorIds(empty, []);

    expect(empty.hasAttribute("data-comment-id")).toBe(false);
    expect(commentAnchorIds(empty)).toEqual([]);
  });
});
