import { afterEach, expect, test } from "bun:test";
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

import {
  failureMarker,
  failureRecord,
  writeFailureRecord,
} from "../test/consumer-scenarios/support/failure-fingerprints";
import { subprocessTestTimeout } from "../test/subprocess";
import {
  baseReplay,
  judgeFailures,
  reportedFailures,
  summarize,
  VERDICTS,
} from "./exploration-base-replay";

const directories: string[] = [];
const temporary = (): string => {
  const directory = mkdtempSync(path.join(tmpdir(), "exploration-base-"));
  directories.push(directory);
  return directory;
};
afterEach(() => {
  for (const directory of directories.splice(0))
    rmSync(directory, { recursive: true, force: true });
});

const REPRO =
  "cd packages/docx-core && PROPERTY_TEST_SEED=-12 PROPERTY_TEST_PATH='3:1:0' PROPERTY_TEST_CASE=0a1b2c3d bun test ./src/ops/__tests__/law.property.test.ts -t 'a law'\\''s case'";

const record = ({
  test,
  repro = REPRO,
  exploration = true,
}: {
  test: string;
  repro?: string;
  exploration?: boolean;
}) =>
  failureRecord(
    {
      ...failureMarker({
        test,
        seed: -12,
        path: "3:1:0",
        repro,
        failure: new Error("law broken"),
      }),
      ...(exploration ? { exploration: true as const } : {}),
    },
    new Error("law broken"),
    { flow: [1, 2] },
  );

test("reads the exploration findings of TypeScript property files, once each", () => {
  const dir = temporary();
  const shard = path.join(dir, "fuzz-exploration-3");
  mkdirSync(shard);
  const law =
    "packages/docx-core/src/ops/__tests__/law.property.test.ts::a law";
  writeFailureRecord(shard, record({ test: law }));
  // The same finding from another shard's artifact.
  writeFailureRecord(dir, record({ test: law }));
  // A pinned or owned failure failed its run: it is not a report.
  writeFailureRecord(
    dir,
    record({ test: `${law} pinned`, exploration: false }),
  );
  // A Rust property finding is filed by its own lane.
  writeFailureRecord(
    dir,
    record({
      test: "crates/docx-kernel::projection laws",
      repro: "PROPTEST_CASES=1 cargo test projection_laws",
    }),
  );
  expect(reportedFailures(dir).map(({ test }) => test)).toEqual([law]);
  expect(reportedFailures(path.join(dir, "missing"))).toEqual([]);
});

test("replays a failure alone on the base, without the test step's reporting environment", () => {
  const replay = baseReplay(
    {
      dir: "packages/docx-core",
      file: "src/ops/__tests__/law.property.test.ts",
      pattern: "a law's case",
      seed: -12,
      path: "3:1:0",
      digest: "0a1b2c3d",
      factor: 4,
    },
    "/base",
    {
      HOME: "/home/runner",
      FOLIO_FUZZ_EXPLORATION: "report",
      FOLIO_FUZZ_EXPLORATION_DIR: "/records",
      FOLIO_FUZZ_EXPLORATION_OWNED: "[]",
      PROPERTY_TEST_SEED: "99",
    },
  );
  expect(replay).toEqual({
    cwd: "/base/packages/docx-core",
    args: [
      "test",
      "./src/ops/__tests__/law.property.test.ts",
      "-t",
      "a law's case",
    ],
    env: {
      HOME: "/home/runner",
      CI: "true",
      PROPERTY_TEST_SEED: "-12",
      PROPERTY_TEST_PATH: "3:1:0",
      PROPERTY_TEST_CASE: "0a1b2c3d",
      PROPERTY_TEST_NUM_RUNS_FACTOR: "4",
    },
  });
});

test("blocks a failure the base passes or cannot replay, and lets one the base shares through", () => {
  const failure = (test: string, repro = REPRO) => ({
    test,
    fingerprint: "f".repeat(16),
    repro,
  });
  const outcomes = {
    "packages/x.test.ts::introduced": { status: "passed", fingerprints: [] },
    "packages/x.test.ts::pre-existing": {
      status: "failed",
      fingerprints: ["f".repeat(16)],
    },
    "packages/x.test.ts::broken": {
      status: "unavailable",
      message: "Replay did not execute tests",
    },
  } as const;
  const failures = Object.keys(outcomes).map((test) => failure(test));
  let current = 0;
  const judged = judgeFailures({
    failures: [...failures, failure("packages/x.test.ts::no line", "bun test")],
    base: "/base",
    env: {},
    run: () => {
      const test = failures[current]?.test as keyof typeof outcomes;
      current += 1;
      return outcomes[test];
    },
  });
  expect(judged.map(({ verdict }) => verdict.type)).toEqual([
    VERDICTS.INTRODUCED,
    VERDICTS.PRE_EXISTING,
    VERDICTS.UNREPLAYABLE,
    VERDICTS.UNREPLAYABLE,
  ]);
  expect(summarize(judged).blocks).toBe(true);
  expect(summarize(judged.slice(1, 2))).toEqual({
    lines: [expect.stringContaining("fails on the base too")],
    blocks: false,
  });
});

test(
  "a real replay that passes on the base blocks; one that runs nothing cannot pass",
  () => {
    const base = temporary();
    mkdirSync(path.join(base, "packages/law"), { recursive: true });
    writeFileSync(
      path.join(base, "packages/law/law.test.ts"),
      'import { test } from "bun:test";\ntest("a law", () => {});\n',
    );
    const replay = (pattern: string) =>
      judgeFailures({
        failures: [
          {
            test: "packages/law/law.test.ts::a law",
            fingerprint: "f".repeat(16),
            repro: `cd packages/law && PROPERTY_TEST_SEED=1 PROPERTY_TEST_PATH='0' bun test ./law.test.ts -t '${pattern}'`,
          },
        ],
        base,
        env: { PATH: process.env["PATH"] ?? "" },
      }).at(0)?.verdict.type;
    expect(replay("a law")).toBe(VERDICTS.INTRODUCED);
    expect(replay("no such test")).toBe(VERDICTS.UNREPLAYABLE);
  },
  subprocessTestTimeout(2),
);
