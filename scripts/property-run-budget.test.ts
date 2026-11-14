import { afterEach, beforeEach, describe, expect, jest, test } from "bun:test";
import fc from "fast-check";

/**
 * Every property run stops under its test's stated budget
 * (test/property-testing.ts, "Bounded runs"): a failure whose shrinking would
 * never end is reported with the counterexample shrunk so far and its replay
 * line before Bun's own timeout, a run cut short of its planned cases fails
 * with the count, and nothing a stopped run started reaches the next test.
 */
import { assertProperty, propertyTestTimeout } from "../test/property-testing";

// These runs fail on purpose: no CI markers, failure records or report mode.
const ENV_KEYS = [
  "CI",
  "PROPERTY_TEST_SEED",
  "PROPERTY_TEST_PATH",
  "FOLIO_PROPERTY_FAILURES_DIR",
  "FOLIO_FUZZ_EXPLORATION",
  "FOLIO_FUZZ_EXPLORATION_DIR",
] as const;
const savedEnv = Object.fromEntries(
  ENV_KEYS.map((key) => [key, process.env[key]]),
);

beforeEach(() => {
  for (const key of ENV_KEYS) Reflect.deleteProperty(process.env, key);
});

afterEach(() => {
  for (const key of ENV_KEYS) {
    const value = savedEnv[key];
    if (value === undefined) Reflect.deleteProperty(process.env, key);
    else process.env[key] = value;
  }
});

/** Shrink steps taken so far: a run still shrinking after its test would add to them. */
let shrinkSteps = 0;

/** Every value shrinks to another value, so a failing run never stops shrinking on its own. */
class EndlessShrink extends fc.Arbitrary<number> {
  generate(): fc.Value<number> {
    return new fc.Value(0, undefined);
  }
  canShrinkWithoutContext(value: unknown): value is number {
    return typeof value === "number";
  }
  shrink(value: number): fc.Stream<fc.Value<number>> {
    shrinkSteps += 1;
    return fc.Stream.of(new fc.Value(value + 1, undefined));
  }
}

const rejection = async (run: Promise<void>): Promise<Error> => {
  const outcome = await run.then(
    () => undefined,
    (error: unknown) => error,
  );
  expect(outcome).toBeInstanceOf(Error);
  return outcome as Error;
};

// Module-level so each test states its budget once, in its timeout argument.
const ENDLESS_BUDGET_MS = propertyTestTimeout(2_000);
const SHORT_BUDGET_MS = propertyTestTimeout(2_000);

describe("property run budgets", () => {
  test(
    "a failure that would shrink forever is reported within its test's budget",
    async () => {
      const budgetMs = ENDLESS_BUDGET_MS;
      const started = Date.now();
      const failure = await rejection(
        assertProperty(
          fc.asyncProperty(new EndlessShrink(), async (value) => {
            await Bun.sleep(5);
            throw new Error(`${String(value)} fails`);
          }),
          { id: "endless shrinking", numRuns: 1 },
        ),
      );
      expect(Date.now() - started).toBeLessThan(budgetMs);
      expect(failure.message).toContain(
        `Shrinking stopped at 90% of the test's ${String(budgetMs)} ms budget`,
      );
      expect(failure.message).toMatch(/Shrunk \d+ time/);
      expect(failure.message).toMatch(
        /Replay: PROPERTY_TEST_SEED=-?\d+ PROPERTY_TEST_PATH=/,
      );
      // It shrank for most of the budget before it was stopped.
      expect(shrinkSteps).toBeGreaterThan(10);
    },
    propertyTestTimeout(2_000),
  );

  test("nothing from the stopped run reaches the next test", async () => {
    const before = shrinkSteps;
    await Bun.sleep(200);
    expect(shrinkSteps).toBe(before);
  });

  test(
    "a property that moves a fake clock is not stopped by the time it fakes",
    () => {
      jest.useFakeTimers();
      try {
        // An hour of fake time per case: a budget read from the faked clock
        // would stop this run after its first case.
        assertProperty(
          fc.property(fc.nat(), () => {
            jest.advanceTimersByTime(3_600_000);
          }),
          { id: "fake clock", numRuns: 50 },
        );
      } finally {
        jest.useRealTimers();
      }
    },
    propertyTestTimeout(2_000),
  );

  test(
    "a run stopped before its planned cases fails and says how many ran",
    async () => {
      const budgetMs = SHORT_BUDGET_MS;
      const failure = await rejection(
        assertProperty(
          fc.asyncProperty(fc.nat(), async () => {
            await Bun.sleep(20);
          }),
          { id: "too many cases for the budget", numRuns: 100_000 },
        ),
      );
      // Stopped short of the test's own timeout, so the count is the failure.
      expect(failure.message).toMatch(
        new RegExp(
          `too many cases for the budget ran \\d+ of \\d+ planned cases before stopping at 95% of the test's ${String(budgetMs)} ms budget`,
        ),
      );
    },
    propertyTestTimeout(2_000),
  );
});
