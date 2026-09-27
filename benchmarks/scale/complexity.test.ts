import { expect, test } from "bun:test";

import { assessComplexity } from "./complexity";

const sizes = [100, 200, 400] as const;

test("linear scaling passes", () => {
  const result = assessComplexity(sizes, [10, 20, 40]);

  expect(result.status).toBe("pass");
  expect(result.timeRatios).toEqual([2, 2]);
  expect(result.normalizedRatios).toEqual([1, 1]);
});

test("quadratic scaling fails with sustained superlinear growth", () => {
  const result = assessComplexity(sizes, [10, 40, 160]);

  expect(result.status).toBe("fail");
  expect(result.normalizedRatios).toEqual([2, 2]);
  expect(result.explanation).toContain("both intervals");
});

test("fixed setup overhead does not trigger the gate", () => {
  const result = assessComplexity(sizes, [1_010, 1_020, 1_040]);

  expect(result.status).toBe("pass");
  expect(result.normalizedRatios[0]).toBeLessThan(1);
  expect(result.normalizedRatios[1]).toBeLessThan(1);
});

test("one noisy interval does not trigger the gate", () => {
  const result = assessComplexity(sizes, [10, 35, 70]);

  expect(result.status).toBe("pass");
  expect(result.normalizedRatios[0]).toBeGreaterThan(result.threshold);
  expect(result.normalizedRatios[1]).toBeLessThanOrEqual(result.threshold);
});
