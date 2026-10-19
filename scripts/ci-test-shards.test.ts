import { afterEach, describe, expect, test } from "bun:test";
import fc from "fast-check";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { propertyConfig, propertyTestTimeout } from "../test/property-testing";
import {
  aggregateTestTimings,
  createGraphqlRequest,
  selectTimingArtifacts,
} from "./aggregate-ci-test-timings";
import { assignTestShards, shardLoads } from "./ci-test-shards";

const temporaryDirectories: string[] = [];

afterEach(() => {
  for (const directory of temporaryDirectories.splice(0)) rmSync(directory, { recursive: true });
});

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

  test("timing artifact discovery distinguishes a failed listing from no data", async () => {
    const failure = new Error("run listing failed");
    expect(selectTimingArtifacts("stella/folio", async () => Promise.reject(failure))).rejects.toBe(
      failure,
    );
    expect(
      await selectTimingArtifacts("stella/folio", async () => ({ workflow_runs: [] })),
    ).toEqual([]);
  });

  test("timing artifact discovery selects complete sets from eligible runs", async () => {
    const fetchJson = async (url: string) => {
      if (url.includes("/workflows/")) {
        return {
          workflow_runs: [
            { event: "pull_request", id: 1 },
            { event: "push", id: 2 },
          ],
        };
      }
      return {
        artifacts: [1, 2, 3, 4].map((id) => ({ id: 20 + id, name: `test-timings-2-${id}` })),
      };
    };
    expect(await selectTimingArtifacts("stella/folio", fetchJson)).toEqual(
      [1, 2, 3, 4].map((id) => ({ artifactId: 20 + id, runId: 2 })),
    );
  });

  test("assignment is deterministic for identical inputs", () => {
    const timings = { a: 10, b: 10, c: 5, d: 5, e: 1 };
    const assigned = () =>
      assignTestShards(Object.keys(timings), timings).map((shard) => Array.from(shard));
    expect(assigned()).toEqual(assigned());
  });

  test("writes large timing contents into a GraphQL request without argv transport", () => {
    const directory = mkdtempSync(path.join(tmpdir(), "folio-timing-request-"));
    temporaryDirectories.push(directory);
    const timingFile = path.join(directory, "timings.json");
    const requestFile = path.join(directory, "request.json");
    const timingBytes = Buffer.from(`${JSON.stringify({ sample: "x".repeat(210_000) })}\n`);
    writeFileSync(timingFile, timingBytes);

    const args = [
      "bun",
      path.join(import.meta.dir, "aggregate-ci-test-timings.ts"),
      "--graphql-request",
      requestFile,
      "--timings",
      timingFile,
      "--repository",
      "stella/folio",
      "--branch",
      "ci/test-timings",
      "--base",
      "0123456789abcdef0123456789abcdef01234567",
    ];
    expect(Math.max(...args.map((argument) => Buffer.byteLength(argument)))).toBeLessThan(4_096);
    const result = Bun.spawnSync(args, { stderr: "pipe" });
    expect(result.exitCode, result.stderr.toString()).toBe(0);

    const request = JSON.parse(readFileSync(requestFile, "utf8"));
    expect(request).toEqual(
      createGraphqlRequest({
        base: "0123456789abcdef0123456789abcdef01234567",
        branch: "ci/test-timings",
        repository: "stella/folio",
        timingFile,
      }),
    );
    expect(
      Buffer.from(request.variables.input.fileChanges.additions[0].contents, "base64"),
    ).toEqual(timingBytes);

    const workflow = readFileSync(
      path.resolve(import.meta.dir, "../.github/workflows/test-timing-refresh.yml"),
      "utf8",
    );
    expect(workflow).toContain('--graphql-request "$request"');
    expect(workflow).not.toContain("--arg contents");
    expect(workflow).not.toContain("< <(");
  });
});
