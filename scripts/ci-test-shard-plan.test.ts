import { expect, setDefaultTimeout, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";

import { scanTestTimeout } from "../test/scan-timeout";
import {
  assignTestShards,
  plannedTestFiles,
  planTestShards,
  SHARD_BUDGET_MS,
  shardLoads,
  type TestTimings,
} from "./ci-test-shard-plan";

setDefaultTimeout(scanTestTimeout(30_000));

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const MINUTE_MS = 60_000;
// A shard's job also installs dependencies and starts a test process per
// suite, and runners vary, so its plan keeps a fifth of the timeout free.
const TIMEOUT_SHARE = 0.8;

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const overTimeout = (loads: readonly number[], timeoutMinutes: number) =>
  loads.flatMap((load, index) =>
    load > timeoutMinutes * MINUTE_MS * TIMEOUT_SHARE
      ? [
          `shard ${index + 1} plans ${load} ms, over ${TIMEOUT_SHARE * 100}% of ${timeoutMinutes} min`,
        ]
      : [],
  );

const shardJobTimeout = (): number => {
  const workflow: unknown = Bun.YAML.parse(
    readFileSync(path.join(REPO_ROOT, ".github/workflows/ci.yml"), "utf8"),
  );
  const jobs = isRecord(workflow) ? workflow["jobs"] : undefined;
  const job = isRecord(jobs) ? jobs["tests"] : undefined;
  const minutes = isRecord(job) ? job["timeout-minutes"] : undefined;
  if (typeof minutes !== "number") throw new Error("tests has no timeout");
  return minutes;
};

test("every planned shard fits within its job's timeout", () => {
  const shards = planTestShards(plannedTestFiles());
  const loads = shardLoads(shards);
  expect(shards.length).toBeGreaterThan(1);
  expect(Math.max(...loads)).toBeLessThanOrEqual(SHARD_BUDGET_MS);
  expect(overTimeout(loads, shardJobTimeout())).toEqual([]);
  expect(overTimeout([SHARD_BUDGET_MS], shardJobTimeout())).toEqual([]);
});

test("a shard planned past the timeout share is reported", () => {
  const timeout = shardJobTimeout();
  const limit = timeout * MINUTE_MS * TIMEOUT_SHARE;
  // Two shards of measurements that only fit one shard per file.
  const timings: TestTimings = {
    "slow.test.ts": limit + 1,
    "other.test.ts": limit / 2,
    "small.test.ts": limit / 4,
  };
  const loads = shardLoads(
    assignTestShards(Object.keys(timings), 2, timings),
    timings,
  );
  expect(overTimeout(loads, timeout)).toEqual([
    `shard 1 plans ${limit + 1} ms, over 80% of ${timeout} min`,
  ]);
});

test("the plan adds a shard once the budget would be exceeded", () => {
  const timings: TestTimings = { a: 40, b: 30, c: 30, d: 20 };
  expect(planTestShards(Object.keys(timings), timings, 120)).toHaveLength(1);
  expect(planTestShards(Object.keys(timings), timings, 70)).toHaveLength(2);
  const three = planTestShards(Object.keys(timings), timings, 50);
  expect(three).toHaveLength(3);
  expect(Math.max(...shardLoads(three, timings))).toBeLessThanOrEqual(50);
  expect(() => planTestShards(Object.keys(timings), timings, 39)).toThrow(
    "on their own: a",
  );
});
