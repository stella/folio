import { describe, expect, test } from "bun:test";

import { deltaDebug } from "./lib/corpus-delta-debug";

const items = Array.from({ length: 64 }, (_, index) => index);

describe("deltaDebug", () => {
  test("shrinks to the items the predicate actually needs", async () => {
    const required = [7, 29, 61];
    const result = await deltaDebug({
      items,
      budget: 1000,
      reproduces: async (kept) => required.every((item) => kept.includes(item)),
    });
    expect([...result.kept].toSorted((left, right) => left - right)).toEqual(required);
    expect(result.exhaustedBudget).toBe(false);
  });

  test("keeps everything when no removal reproduces", async () => {
    const result = await deltaDebug({
      items,
      budget: 200,
      reproduces: async (kept) => kept.length === items.length,
    });
    expect(result.kept).toHaveLength(items.length);
  });

  test("stops at the budget and reports it", async () => {
    let calls = 0;
    const result = await deltaDebug({
      items,
      budget: 3,
      reproduces: async (kept) => {
        calls += 1;
        return kept.includes(0);
      },
    });
    expect(calls).toBe(3);
    expect(result.evaluations).toBe(3);
    expect(result.exhaustedBudget).toBe(true);
  });

  test("a result is 1-minimal: removing any single survivor stops reproducing", async () => {
    const required = new Set([3, 11]);
    const reproduces = async (kept: readonly number[]) =>
      [...required].every((item) => kept.includes(item));
    const { kept } = await deltaDebug({ items, budget: 1000, reproduces });
    for (const survivor of kept) {
      // oxlint-disable-next-line no-await-in-loop -- one check per survivor, by definition sequential
      expect(await reproduces(kept.filter((item) => item !== survivor))).toBe(false);
    }
  });
});
