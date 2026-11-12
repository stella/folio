import { expect, setDefaultTimeout, test } from "bun:test";
import { existsSync, readFileSync } from "node:fs";
import path from "node:path";

import { scanTestTimeout } from "../test/scan-timeout";

setDefaultTimeout(scanTestTimeout(30_000));

const REPO_ROOT = path.resolve(import.meta.dir, "..");

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

// A job without its own limit runs for the platform default of six hours, so
// a stalled step holds a runner and the checks waiting on it. Every job states
// the limit its normal duration needs, with headroom.
//
// A job that calls a reusable workflow cannot take timeout-minutes. A callee
// in this repository is a workflow file of its own and is checked like every
// other file. A callee in another repository owns its limits there.
const LOCAL_CALLEE = /^\.\/\.github\/workflows\/[^@]+\.ya?ml$/u;
const REMOTE_CALLEE =
  /^[\w.-]+\/[\w.-]+\/\.github\/workflows\/[^@]+@[0-9a-f]{40}$/u;

const jobsWithoutTimeouts = (
  label: string,
  workflow: unknown,
  exists: (callee: string) => boolean,
): string[] => {
  if (!isRecord(workflow) || !isRecord(workflow["jobs"]))
    return [`${label}: no jobs mapping`];
  return Object.entries(workflow["jobs"]).flatMap(([name, job]) => {
    if (!isRecord(job)) return [`${label}: ${name} is not a mapping`];
    const uses = job["uses"];
    if (typeof uses === "string") {
      if (LOCAL_CALLEE.test(uses))
        return exists(uses)
          ? []
          : [`${label}: ${name} calls missing workflow ${uses}`];
      return REMOTE_CALLEE.test(uses)
        ? []
        : [`${label}: ${name} calls ${uses}, neither local nor pinned`];
    }
    const minutes = job["timeout-minutes"];
    return typeof minutes === "number" &&
      Number.isInteger(minutes) &&
      minutes > 0 &&
      minutes < 360
      ? []
      : [`${label}: ${name} declares no timeout-minutes below 360`];
  });
};

const committedWorkflows = (): [string, unknown][] =>
  [
    ...new Bun.Glob(".github/workflows/*.{yml,yaml}").scanSync({
      cwd: REPO_ROOT,
    }),
  ]
    .toSorted()
    .map((file) => [
      file,
      Bun.YAML.parse(readFileSync(path.join(REPO_ROOT, file), "utf8")),
    ]);

const existsInRepo = (callee: string) =>
  existsSync(path.join(REPO_ROOT, callee));

test("every committed job declares its own timeout", () => {
  const workflows = committedWorkflows();
  expect(workflows.length).toBeGreaterThan(0);
  expect(
    workflows.flatMap(([file, workflow]) =>
      jobsWithoutTimeouts(file, workflow, existsInRepo),
    ),
  ).toEqual([]);
});

test("a job whose timeout is removed is reported", () => {
  const [file, workflow] =
    committedWorkflows().find(([name]) => name.endsWith("/ci.yml")) ?? [];
  if (file === undefined || !isRecord(workflow) || !isRecord(workflow["jobs"]))
    throw new Error("ci.yml has no jobs");
  const broken = structuredClone(workflow);
  const jobs = broken["jobs"];
  if (!isRecord(jobs) || !isRecord(jobs["packaged-consumer"]))
    throw new Error("ci.yml has no packaged-consumer job");
  delete jobs["packaged-consumer"]["timeout-minutes"];
  expect(jobsWithoutTimeouts(file, broken, existsInRepo)).toEqual([
    `${file}: packaged-consumer declares no timeout-minutes below 360`,
  ]);
});

test("reusable calls are accepted only for a present or pinned callee", () => {
  const sha = "c6d66dfa2a5ad1f96a21b819fe456fecdd3d6f6d";
  const calls = (uses: string) => ({ jobs: { call: { uses } } });
  expect(
    jobsWithoutTimeouts(
      "fixture",
      calls("./.github/workflows/fuzz-failure-report.yml"),
      existsInRepo,
    ),
  ).toEqual([]);
  expect(
    jobsWithoutTimeouts(
      "fixture",
      calls(`stella/.github/.github/workflows/cla.yml@${sha}`),
      existsInRepo,
    ),
  ).toEqual([]);
  expect(
    jobsWithoutTimeouts(
      "fixture",
      calls("./.github/workflows/absent.yml"),
      existsInRepo,
    ),
  ).toHaveLength(1);
  expect(
    jobsWithoutTimeouts(
      "fixture",
      calls("stella/.github/.github/workflows/cla.yml@main"),
      existsInRepo,
    ),
  ).toHaveLength(1);
  for (const minutes of [undefined, 0, 360, "30", 1.5]) {
    expect(
      jobsWithoutTimeouts(
        "fixture",
        { jobs: { run: { "timeout-minutes": minutes } } },
        existsInRepo,
      ),
    ).toHaveLength(1);
  }
});
