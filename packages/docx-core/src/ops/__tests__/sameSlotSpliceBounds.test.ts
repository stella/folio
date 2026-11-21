import { expect, test } from "bun:test";

import { changedListBounds, spliceItems } from "../listSplice";

const observed = (width: number) => {
  let reads = 0;
  const source = new Proxy(
    Array.from({ length: width }, (_, index) => index),
    {
      get(target, key, receiver) {
        if (typeof key === "string" && /^\d+$/u.test(key)) reads++;
        return Reflect.get(target, key, receiver);
      },
    },
  );
  return {
    source,
    reads: () => reads,
    reset: () => {
      reads = 0;
    },
  };
};
const replace = (items: number[], index: number, id: number) =>
  spliceItems({ items, start: index, count: 1, replacement: [id] });

test.each([1024, 4096])(
  "recover same-slot ancestor and sibling bounds in both directions without reading %i slots",
  (width) => {
    const { source, reads, reset } = observed(width);
    const at = width / 2;
    const one = replace(source, at, -1);
    const sibling = replace(source, at, -2);
    const two = replace(one, at, -3);
    let deep = two;
    for (let step = 0; step < 16; step++) deep = replace(deep, at, -4);

    let siblingReads = 0;
    Object.defineProperty(sibling, 0, {
      get: () => {
        siblingReads++;
        return 0;
      },
    });
    reset();
    const changed = { start: at, previousEnd: at + 1, end: at + 1 };
    for (const [previous, current] of [
      [source, two],
      [two, source],
      [sibling, two],
      [two, sibling],
      [sibling, deep],
      [deep, sibling],
    ] as const) {
      expect(changedListBounds(previous, current)).toEqual(changed);
      expect(reads()).toBe(0);
      expect(siblingReads).toBe(0);
    }
    expect(source[at]).toBe(at);
    expect(one[at]).toBe(-1);
    expect(sibling[at]).toBe(-2);
    expect(two[at]).toBe(-3);
    expect(deep[at]).toBe(-4);
  },
);

test("unknown histories and edits that shift or change a different slot retain the fallback", () => {
  const width = 1024;
  const { source, reads, reset } = observed(width);
  const one = replace(source, 20, -1);
  const disjoint = replace(one, 40, -1);
  const adjacent = replace(one, 21, -1);
  const growth = spliceItems({
    items: one,
    start: 34,
    count: 0,
    replacement: [-1],
  });
  const restoredLength = spliceItems({
    items: growth,
    start: 34,
    count: 1,
    replacement: [],
  });
  const unknown = replace(source.slice(), 20, -1);

  for (const [target, expected] of [
    [unknown, { start: 20, previousEnd: 21, end: 21 }],
    [disjoint, { start: 20, previousEnd: 41, end: 41 }],
    [adjacent, { start: 20, previousEnd: 22, end: 22 }],
    [growth, { start: 20, previousEnd: 34, end: 35 }],
    [restoredLength, { start: 20, previousEnd: 21, end: 21 }],
  ] as const) {
    reset();
    expect(changedListBounds(source, target)).toEqual(expected);
    expect(reads()).toBeGreaterThan(990);
  }
});

test("every changed slot stays inside the bounds across forks and repeated values", () => {
  const source = Array.from({ length: 32 }, (_, index) => index);
  const versions = [source];
  for (let index = 0; index < source.length; index++) {
    const first = replace(source, index, -1);
    const restored = replace(first, index, index);
    const repeated = replace(first, index, -1);
    const chained = replace(repeated, index, -2);
    const disjoint = replace(first, (index + 13) % source.length, -1);
    versions.push(first, restored, repeated, chained, disjoint);
  }
  for (const before of versions) {
    for (const after of versions) {
      const bounds = changedListBounds(before, after);
      expect(bounds.previousEnd).toBe(bounds.end);
      expect(before.slice(0, bounds.start)).toEqual(
        after.slice(0, bounds.start),
      );
      expect(before.slice(bounds.previousEnd)).toEqual(after.slice(bounds.end));
    }
  }
});
