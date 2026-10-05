import { expect, test } from "bun:test";
import fc from "fast-check";
import { BROWSER_FUZZ_BUDGET, checkWithBoundedShrink } from "../test/bounded-async-fuzz";

test("browser fuzz budgets leave time for setup and durable reporting", () => {
  expect(BROWSER_FUZZ_BUDGET.discoveryMs + BROWSER_FUZZ_BUDGET.shrinkMs).toBeLessThan(
    BROWSER_FUZZ_BUDGET.testMs,
  );
});

test("a stalled shrink retains the first assertion and seeded witness", async () => {
  const original = new TypeError("Original invariant failed");
  const retained: { seed: number; path: string; value: number[]; error: unknown }[] = [];
  const pending: (() => void)[] = [];
  let evaluations = 0;
  const result = await checkWithBoundedShrink({
    arbitrary: fc.array(fc.integer(), { minLength: 2, maxLength: 8 }),
    seed: 347,
    numRuns: 20,
    discoveryMs: 1_000,
    shrinkMs: 20,
    evaluate: async () => {
      evaluations++;
      if (evaluations === 1) throw original;
      await new Promise<void>((resolve) => pending.push(resolve));
    },
    onFirstFailure: async (failure) => {
      expect(evaluations).toBe(1);
      retained.push(failure);
    },
  });
  for (const resolve of pending) resolve();
  expect(retained).toHaveLength(1);
  expect(retained.at(0)?.seed).toBe(347);
  expect(retained.at(0)?.error).toBe(original);
  expect(retained.at(0)?.value).toEqual(result.discovery.counterexample?.at(0));
  expect(retained.at(0)?.path).toBe(result.discovery.counterexamplePath);
  expect(result.discovery.failed).toBe(true);
  expect(result.shrinking?.failed).toBe(true);
  expect(result.shrinking?.interrupted).toBe(true);
  expect(result.verdict).toBe(result.discovery);
  expect(result.verdict.errorInstance).toBe(original);
  expect(evaluations).toBe(2);
});

test("a replay mismatch is a harness failure after retaining the original finding", async () => {
  let offset = 0;
  let retained = 0;
  const result = checkWithBoundedShrink({
    arbitrary: fc.integer().map((value) => value + offset),
    seed: 347,
    numRuns: 1,
    discoveryMs: 1_000,
    shrinkMs: 1_000,
    evaluate: async () => {
      throw new TypeError("Original assertion");
    },
    onFirstFailure: async () => {
      retained++;
      offset = 1;
    },
  });
  await expect(result).rejects.toThrow(
    "The seeded shrink replay did not start from the retained failure.",
  );
  expect(retained).toBe(1);
});

test("non-root initial paths preserve discovery and minimized replay witnesses", async () => {
  const arbitrary = fc.array(fc.integer(), { minLength: 1, maxLength: 8 });
  const seed = 347;
  const baseline = await fc.check(
    fc.asyncProperty(arbitrary, async () => false),
    { seed, numRuns: 1 },
  );
  if (baseline.counterexamplePath === null) throw new TypeError("Missing generated replay path");
  const result = await checkWithBoundedShrink({
    arbitrary,
    seed,
    path: baseline.counterexamplePath,
    numRuns: 1,
    discoveryMs: 1_000,
    shrinkMs: 1_000,
    evaluate: async () => {
      throw new TypeError("Invariant failed");
    },
    onFirstFailure: () => Promise.resolve(),
  });
  for (const verdict of [result.discovery, result.shrinking]) {
    if (!verdict || verdict.counterexamplePath === null)
      throw new TypeError("Missing replay witness");
    expect(
      fc.sample(arbitrary, { seed, path: verdict.counterexamplePath, numRuns: 1 }).at(0),
    ).toEqual(verdict.counterexample?.at(0));
  }
});

test("generated failures are saved once before shrinking without replaying the first witness", async () => {
  for (const seed of [11, 29, 47, 347]) {
    const evaluated: number[][] = [];
    let saved = 0;
    const result = await checkWithBoundedShrink({
      arbitrary: fc.array(fc.integer(), { minLength: 1, maxLength: 8 }),
      seed,
      numRuns: 10,
      discoveryMs: 1_000,
      shrinkMs: 1_000,
      evaluate: async (value) => {
        evaluated.push([...value]);
        throw new TypeError("Invariant failed");
      },
      onFirstFailure: async () => {
        saved++;
        expect(evaluated).toHaveLength(1);
      },
    });
    expect(saved).toBe(1);
    expect(result.discovery.failed).toBe(true);
    expect(result.shrinking?.failed).toBe(true);
    expect(
      evaluated.filter((value) => JSON.stringify(value) === JSON.stringify(evaluated.at(0))),
    ).toHaveLength(1);
  }
});
