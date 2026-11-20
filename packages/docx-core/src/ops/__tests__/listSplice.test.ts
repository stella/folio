import { expect, test } from "bun:test";

import type { Paragraph } from "../../model/document";
import { hasMultipleParagraphPropertyChanges } from "../contract";
import { changedListBounds, spliceItems } from "../listSplice";

test("splice boundaries apply only to their exact source", () => {
  const source = ["a", "b", "c"];
  const result = spliceItems({
    items: source,
    start: 1,
    count: 1,
    replacement: ["x", "y"],
  });
  expect(changedListBounds(source, result)).toEqual(
    expect.objectContaining({ start: 1, previousEnd: 2, end: 3 }),
  );
  expect(changedListBounds(["z", "b", "c"], result)).toEqual({
    start: 0,
    previousEnd: 2,
    end: 3,
  });
});

test.each([16, 64])(
  "property checks visit only replacement paragraphs across %i edits",
  (size) => {
    let reads = 0;
    const paragraph: Paragraph = {
      type: "paragraph",
      content: [],
      get propertyChanges() {
        reads++;
        return undefined;
      },
    };
    let content = Array.from({ length: size }, () => paragraph);
    const check = () =>
      hasMultipleParagraphPropertyChanges({
        package: { document: { content } },
      });
    expect(check()).toBe(false);
    expect(reads).toBe(size);
    for (let i = 0; i < size; i++) {
      content = spliceItems({
        items: content,
        start: 0,
        count: 0,
        replacement: [paragraph],
      });
      expect(check()).toBe(false);
    }
    expect(reads).toBe(2 * size);
  },
);
