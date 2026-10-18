import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import { propertyConfig, propertyTestTimeout } from "../test/property-testing";
import { aggregateTestTimings } from "./aggregate-ci-test-timings";
import { assignTestShards, shardLoads } from "./ci-test-shards";

describe("measured test shard balancing", () => {
  test(
    "longest-first packing stays within 1.25x mean when no file exceeds the mean",
    () => {
      fc.assert(
        fc.property(
          fc.array(fc.integer({ min: 1, max: 1_000 }), { minLength: 100, maxLength: 100 }),
          (weights) => {
            const timings = Object.fromEntries(
              weights.map((weight, index) => [`test-${index}`, weight]),
            );
            const mean = weights.reduce((sum, weight) => sum + weight, 0) / 4;
            fc.pre(Math.max(...weights) <= mean);
            const loads = shardLoads(assignTestShards(Object.keys(timings), timings), timings);
            expect(Math.max(...loads)).toBeLessThanOrEqual(mean * 1.25);
          },
        ),
        propertyConfig(),
      );
    },
    propertyTestTimeout(5_000),
  );

  test("unmeasured files use the median measured weight", () => {
    const timings = { a: 10, b: 30, c: 50 };
    expect(shardLoads(assignTestShards(["a", "b", "c", "new"], timings), timings)).toEqual([
      50, 30, 30, 10,
    ]);
  });

  test("aggregation takes per-run medians, follows discovery, and is byte-stable", () => {
    const artifacts = [
      { runId: "one", timings: { kept: 10, vanished: 1 } },
      { runId: "two", timings: { kept: 30 } },
      { runId: "three", timings: { kept: 20 } },
    ];
    const first = aggregateTestTimings(artifacts, ["new", "kept"]);
    expect(first).toBe(aggregateTestTimings(artifacts.toReversed(), ["kept", "new"]));
    expect(JSON.parse(first)).toEqual({ kept: 20, new: 15 });
    expect(first).not.toContain("vanished");
  });

  test("assignment is deterministic for identical inputs", () => {
    const timings = { a: 10, b: 10, c: 5, d: 5, e: 1 };
    const assigned = () =>
      assignTestShards(Object.keys(timings), timings).map((shard) => Array.from(shard));
    expect(assigned()).toEqual(assigned());
  });
});
