import { panic } from "better-result";
import { afterEach, expect, test } from "bun:test";
import {
  mkdirSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

import {
  expectExited,
  runChild,
  subprocessTestTimeout,
} from "../test/subprocess";
import {
  aggregateTestTimings,
  selectTimingArtifacts,
} from "./aggregate-ci-test-timings";
import {
  createFullSuiteTimingArtifact,
  fullSuiteTimingSets,
} from "./ci-test-timing-artifact";

const RUN_ID = "44002681576";
const OTHER_RUN = "43997120235";
const complete = (shards: number) =>
  Array.from({ length: shards }, (_, index) => ({
    runId: RUN_ID,
    artifact: createFullSuiteTimingArtifact({
      runId: RUN_ID,
      shard: index + 1,
      shards,
      timings: {
        [`packages/core/src/part-${index + 1}.test.ts`]:
          index === 0 ? 683_233 : index + 1,
      },
    }),
  }));

for (const shards of [4, 6]) {
  test(`a complete ${shards}-shard run retains every measured millisecond`, () => {
    const inputs = complete(shards);
    const result = fullSuiteTimingSets(inputs.toReversed());
    expect(result.legacyRuns).toEqual([]);
    expect(result.artifacts).toHaveLength(shards);
    const files = inputs.flatMap(({ artifact }) =>
      Object.keys(artifact.timings),
    );
    expect(JSON.parse(aggregateTestTimings(result.artifacts, files))).toEqual(
      Object.assign({}, ...inputs.map(({ artifact }) => artifact.timings)),
    );
  });
  test(`discovery includes ${shards}-shard candidates without guessing completeness`, async () => {
    const selected = await selectTimingArtifacts("stella/folio", async (url) =>
      url.includes("/workflows/")
        ? { workflow_runs: [{ event: "merge_group", id: Number(RUN_ID) }] }
        : {
            artifacts: complete(shards).map(({ artifact }) => ({
              id: artifact.shard,
              name: `test-timings-${RUN_ID}-${artifact.shard}`,
            })),
          },
    );
    expect(selected).toHaveLength(shards);
  });
  for (let missing = 1; missing <= shards; missing++)
    test(`${shards}-shard provenance refuses missing member ${missing}`, () => {
      const inputs = complete(shards).filter(
        ({ artifact }) => artifact.shard !== missing,
      );
      expect(() => fullSuiteTimingSets(inputs)).toThrow(
        "missing planned shard measurements",
      );
    });
}

test("duplicate shard, inconsistent plan and duplicate file ownership are distinct failures", () => {
  const inputs = complete(6);
  const first = inputs.at(0) ?? panic("A fixture shard is missing");
  expect(() => fullSuiteTimingSets([...inputs, first])).toThrow(
    "repeats an executed shard",
  );
  const inconsistent = inputs.map((entry, index) =>
    index === 1
      ? { runId: entry.runId, artifact: { ...entry.artifact, shards: 4 } }
      : entry,
  );
  expect(() => fullSuiteTimingSets(inconsistent)).toThrow(
    "inconsistent planned shard counts",
  );
  const duplicateFile = inputs.map((entry, index) =>
    index === 1
      ? {
          runId: entry.runId,
          artifact: { ...entry.artifact, timings: first.artifact.timings },
        }
      : entry,
  );
  expect(() => fullSuiteTimingSets(duplicateFile)).toThrow(
    "assigns one file to multiple shards",
  );
});

test("the manifest run identity is authoritative", () => {
  const inputs = complete(4).map(({ artifact }) => ({
    runId: OTHER_RUN,
    artifact,
  }));
  expect(() => fullSuiteTimingSets(inputs)).toThrow(
    "different run than its download manifest",
  );
});

test("local execution has an explicit origin and cannot refresh GitHub measurements", () => {
  const artifact = createFullSuiteTimingArtifact({
    runId: undefined,
    shard: 1,
    shards: 1,
    timings: { "sample.test.ts": 10 },
  });
  expect(artifact.origin).toEqual({ type: "local" });
  expect(() => fullSuiteTimingSets([{ runId: RUN_ID, artifact }])).toThrow(
    "Local full-suite timings cannot refresh CI",
  );
  expect(() =>
    createFullSuiteTimingArtifact({
      runId: "local",
      shard: 1,
      shards: 1,
      timings: {},
    }),
  ).toThrow("GitHub run id");
});

test("invalid shard facts and malformed measurements do not masquerade as old provenance", () => {
  const first = complete(4).at(0) ?? panic("A fixture shard is missing");
  for (const change of [
    { shard: 0 },
    { shard: 5 },
    { shards: 0 },
    { shard: 1.5 },
  ])
    expect(() =>
      fullSuiteTimingSets([
        { runId: RUN_ID, artifact: { ...first.artifact, ...change } },
      ]),
    ).toThrow("valid executed shard");
  for (const duration of [0, -1, Number.NaN, Number.POSITIVE_INFINITY, "10"])
    expect(() =>
      fullSuiteTimingSets([
        {
          runId: RUN_ID,
          artifact: {
            ...first.artifact,
            timings: { "sample.test.ts": duration },
          },
        },
      ]),
    ).toThrow("positive finite milliseconds");
  expect(() =>
    fullSuiteTimingSets([
      { runId: RUN_ID, artifact: { runId: RUN_ID, shard: 1, timings: {} } },
    ]),
  ).toThrow("Unsupported full-suite timing artifact version");
});

test("malformed count-less JSON is not mistaken for a legacy timing artifact", () => {
  expect(() =>
    fullSuiteTimingSets([{ runId: RUN_ID, artifact: { runId: RUN_ID } }]),
  ).toThrow("missing file measurements");
  expect(() =>
    fullSuiteTimingSets([
      {
        runId: RUN_ID,
        artifact: { runId: RUN_ID, timings: { "sample.test.ts": -1 } },
      },
    ]),
  ).toThrow("positive finite milliseconds");
});

test("legacy provenance is reported, never treated as a complete run", () => {
  const legacy = {
    runId: OTHER_RUN,
    artifact: { runId: OTHER_RUN, timings: { "sample.test.ts": 10 } },
  };
  const result = fullSuiteTimingSets([...complete(6), legacy]);
  expect(result.artifacts).toHaveLength(6);
  expect(result.legacyRuns).toEqual([OTHER_RUN]);
  expect(fullSuiteTimingSets([legacy]).artifacts).toEqual([]);
  expect(() =>
    fullSuiteTimingSets([
      ...complete(4),
      {
        ...legacy,
        runId: RUN_ID,
        artifact: { ...legacy.artifact, runId: RUN_ID },
      },
    ]),
  ).toThrow("mixes legacy and proven");
});

const directories: string[] = [];
afterEach(() => {
  for (const directory of directories.splice(0))
    rmSync(directory, { recursive: true, force: true });
});

for (const origin of ["incomplete", "legacy-only"] as const)
  test(
    `aggregation preserves the committed file for ${origin} input`,
    () => {
      const directory = mkdtempSync(
        path.join(tmpdir(), "folio-timing-provenance-"),
      );
      directories.push(directory);
      const root = path.join(directory, "artifacts");
      const output = path.join(directory, "timings.json");
      const original = '{"standing.test.ts":683233}\n';
      writeFileSync(output, original);
      const values =
        origin === "incomplete"
          ? complete(6)
              .slice(0, 5)
              .map(({ artifact }) => artifact)
          : [{ runId: RUN_ID, timings: { "sample.test.ts": 10 } }];
      const inputs = values.map((artifact, index) => {
        const folder = path.join(root, RUN_ID, String(index + 100));
        mkdirSync(folder, { recursive: true });
        const file = path.join(folder, `test-timings-${index + 1}.json`);
        writeFileSync(file, JSON.stringify(artifact));
        return file;
      });
      const result = expectExited(
        runChild([
          "bun",
          path.join(import.meta.dir, "aggregate-ci-test-timings.ts"),
          "--source",
          "tests",
          "--output",
          output,
          "--artifact-root",
          root,
          ...inputs,
        ]),
      );
      expect(result.exitCode).toBe(origin === "incomplete" ? 1 : 0);
      expect(result.stderr).toContain(
        origin === "incomplete"
          ? "missing planned shard measurements"
          : "No proven full-suite timing sets",
      );
      expect(readFileSync(output, "utf8")).toBe(original);
    },
    subprocessTestTimeout(),
  );

test(
  "extended properties retain their independently valid partial timing set",
  () => {
    const directory = mkdtempSync(
      path.join(tmpdir(), "folio-property-timings-"),
    );
    directories.push(directory);
    const input = path.join(directory, "property-timings-1.json");
    const output = path.join(directory, "property-timings.json");
    const file = "packages/docx-core/src/ops/__tests__/apply.property.test.ts";
    writeFileSync(
      input,
      JSON.stringify({ runId: RUN_ID, timings: { [file]: 659_635 } }),
    );
    const result = expectExited(
      runChild([
        "bun",
        path.join(import.meta.dir, "aggregate-ci-test-timings.ts"),
        "--source",
        "properties",
        "--output",
        output,
        input,
      ]),
    );
    expect(result.exitCode, result.stderr).toBe(0);
    expect(JSON.parse(readFileSync(output, "utf8"))).toEqual({
      [file]: 659_635,
    });
  },
  subprocessTestTimeout(),
);
