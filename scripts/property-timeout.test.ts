import { afterEach, expect, test } from "bun:test";
import { propertyConfig, propertyTestTimeout as sharedTimeout } from "../test/property-testing";
import { NUM_RUNS_FACTOR_ENV } from "../test/property-run-factor";
import { propertyTestTimeout } from "../test/property-timeout";

const originalFactor = process.env[NUM_RUNS_FACTOR_ENV];
afterEach(() => {
  if (originalFactor === undefined) Reflect.deleteProperty(process.env, NUM_RUNS_FACTOR_ENV);
  else process.env[NUM_RUNS_FACTOR_ENV] = originalFactor;
});

test.each([1, 5, 10, 1.5])("stated budgets and case counts scale by %s", (factor) => {
  process.env[NUM_RUNS_FACTOR_ENV] = String(factor);
  expect(sharedTimeout).toBe(propertyTestTimeout);
  for (const base of [0.5, 1, 15_000, 30_000]) {
    expect(propertyTestTimeout(base)).toBe(Math.ceil(base * factor));
  }
  expect(propertyConfig({ numRuns: 100 }).numRuns).toBe(Math.ceil(100 * factor));
});

test.each([Number.NaN, Number.POSITIVE_INFINITY, Number.NEGATIVE_INFINITY, 0, -1])(
  "invalid base budget %s fails before test registration",
  (base) => expect(() => propertyTestTimeout(base)).toThrow("finite and positive"),
);

test("overflowing scaled budgets fail before test registration", () => {
  process.env[NUM_RUNS_FACTOR_ENV] = "10";
  expect(() => propertyTestTimeout(Number.MAX_VALUE)).toThrow("must be finite");
});
