import { expect, spyOn, test } from "bun:test";

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
  "inverse splice checks leave the unchanged prefix unread at length %i",
  (size) => {
    let reads = 0;
    const source = Array.from({ length: size }, (_, index) => index);
    Object.defineProperty(source, 0, {
      get: () => {
        reads++;
        return 0;
      },
    });
    const result = spliceItems({
      items: source,
      start: size - 1,
      count: 1,
      replacement: [-1, -2],
    });
    const before = reads;
    expect(changedListBounds(result, source)).toEqual({
      start: size - 1,
      previousEnd: size + 1,
      end: size,
    });
    expect(reads).toBe(before);
  },
);

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

test("large splices preserve both boundaries without using replacement elements as arguments", () => {
  const size = 150_000;
  const replacement = Array.from({ length: size }, (_, index) => index);
  const source = [-1, -2, -3];
  const native = spyOn(source, "toSpliced");
  try {
    const full = spliceItems({
      items: source,
      start: 0,
      count: source.length,
      replacement,
    });
    expect(full).toEqual(replacement);
    expect(full).not.toBe(replacement);
    expect(changedListBounds(source, full)).toMatchObject({
      start: 0,
      previousEnd: source.length,
      end: size,
    });

    const middle = spliceItems({
      items: source,
      start: 1,
      count: 1,
      replacement,
    });
    expect(middle.at(0)).toBe(-1);
    expect(middle.slice(1, -1)).toEqual(replacement);
    expect(middle.at(-1)).toBe(-3);
    expect(changedListBounds(middle, source)).toEqual({
      start: 1,
      previousEnd: size + 1,
      end: 2,
    });
    expect(native.mock.calls.map((args) => args.length)).toEqual([2]);
  } finally {
    native.mockRestore();
  }
});

test("large splices keep native numeric, sparse, and removed-element behavior", () => {
  const replacement = Array.from({ length: 1_025 }, (_, index) => index);
  const source = [-1, -2, -3, -4];
  Reflect.deleteProperty(source, 0);
  Object.defineProperty(source, 2, {
    get: () => {
      throw new Error("Splicing must not read an element it removes.");
    },
  });
  const next = spliceItems({
    items: source,
    start: -2.9,
    count: 1.9,
    replacement,
  });
  expect(Object.hasOwn(next, 0)).toBe(true);
  expect(next.slice(0, 2)).toEqual([undefined, -2]);
  expect(next.slice(2, -1)).toEqual(replacement);
  expect(next.at(-1)).toBe(-4);
});
