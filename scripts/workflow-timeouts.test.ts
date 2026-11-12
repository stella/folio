import { expect, setDefaultTimeout, test } from "bun:test";
import { existsSync, readFileSync } from "node:fs";
import path from "node:path";

import { scanTestTimeout } from "../test/scan-timeout";
import {
  aggregateWorkflowDurations,
  type CheckRun,
  committedWorkflows as workflowsByFile,
  limitsUnderMeasuredDurations,
  minimumMinutes,
  p95,
  parseWorkflowRuns,
  readWorkflowDurations,
  WORKFLOW_DURATIONS_FILE,
} from "./workflow-durations";

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

const committedWorkflows = (): [string, unknown][] => [...workflowsByFile()];

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

// A limit a healthy run can reach fails unrelated changes with every check
// passing. Each declared limit keeps a fifth of itself free over the measured
// 95th percentile of its job or step (scripts/workflow-durations.json, rewritten
// by .github/workflows/test-timing-refresh.yml).
const REPLAY_STEP = "Run only standing browser input regressions";

test("every limit keeps a fifth free over its measured p95", () => {
  const durations = readWorkflowDurations();
  expect(limitsUnderMeasuredDurations(workflowsByFile(), durations)).toEqual(
    [],
  );
  const pinned =
    durations[".github/workflows/ci.yml"]?.["browser-input-pinned"];
  expect(pinned?.job?.samples).toBeGreaterThan(0);
  expect(pinned?.steps?.[REPLAY_STEP]?.samples).toBeGreaterThan(0);
});

test("a job or step limit under its measured p95 share is reported", () => {
  const file = ".github/workflows/ci.yml";
  const durations = readWorkflowDurations();
  const jobs = durations[file];
  const stepP95 =
    jobs?.["browser-input-pinned"]?.steps?.[REPLAY_STEP]?.p95Seconds;
  const jobP95 = jobs?.["packaged-consumer"]?.job?.p95Seconds;
  if (stepP95 === undefined || jobP95 === undefined)
    throw new Error(
      `${WORKFLOW_DURATIONS_FILE} lacks the measured ci.yml entries`,
    );
  const broken = structuredClone(workflowsByFile().get(file));
  const brokenJobs = isRecord(broken) ? broken["jobs"] : undefined;
  const pinned = isRecord(brokenJobs)
    ? brokenJobs["browser-input-pinned"]
    : undefined;
  const consumer = isRecord(brokenJobs)
    ? brokenJobs["packaged-consumer"]
    : undefined;
  const steps = isRecord(pinned) ? pinned["steps"] : undefined;
  const replay = Array.isArray(steps)
    ? steps.find((step) => isRecord(step) && step["name"] === REPLAY_STEP)
    : undefined;
  if (!isRecord(replay) || !isRecord(consumer))
    throw new Error("ci.yml lacks the pinned replay step or packaged-consumer");
  const stepMinutes = minimumMinutes(stepP95) - 1;
  const jobMinutes = minimumMinutes(jobP95) - 1;
  replay["timeout-minutes"] = stepMinutes;
  consumer["timeout-minutes"] = jobMinutes;
  expect(
    limitsUnderMeasuredDurations(new Map([[file, broken]]), durations),
  ).toEqual([
    `${file}: browser-input-pinned step "${REPLAY_STEP}" allows ${stepMinutes} min, but its measured p95 of ${stepP95} s needs at least ${stepMinutes + 1} min`,
    `${file}: packaged-consumer allows ${jobMinutes} min, but its measured p95 of ${jobP95} s needs at least ${jobMinutes + 1} min`,
  ]);
});

test("a step limit needs a literal name to be measured by", () => {
  const workflow = {
    jobs: {
      run: {
        "timeout-minutes": 5,
        steps: [
          { name: "Named", run: "true", "timeout-minutes": 2 },
          { run: "true", "timeout-minutes": 2 },
          {
            name: "Shard ${{ matrix.shard }}",
            run: "true",
            "timeout-minutes": 2,
          },
        ],
      },
    },
  };
  expect(
    limitsUnderMeasuredDurations(new Map([["fixture", workflow]]), {}),
  ).toEqual([
    "fixture: run step 2 declares timeout-minutes without a literal name to measure it by",
    "fixture: run step 3 declares timeout-minutes without a literal name to measure it by",
  ]);
});

test("the rule takes the nearest-rank p95 and whole minutes", () => {
  expect(p95(Array.from({ length: 20 }, (_, index) => index + 1))).toBe(19);
  expect(p95(Array.from({ length: 100 }, (_, index) => 100 - index))).toBe(95);
  expect(p95([7])).toBe(7);
  expect(minimumMinutes(240)).toBe(5);
  expect(minimumMinutes(241)).toBe(6);
});

const at = (seconds: number) =>
  new Date(Date.UTC(2026, 0, 1, 0, 0, seconds)).toISOString();

const checkRun = (
  name: string,
  seconds: number,
  conclusion = "SUCCESS",
  steps: CheckRun["steps"] = [],
): CheckRun => ({
  id: `${name}-${seconds}`,
  name,
  conclusion,
  startedAt: at(0),
  completedAt: at(seconds),
  steps,
});

const step = (name: string, seconds: number, conclusion = "SUCCESS") => ({
  name,
  conclusion,
  startedAt: at(0),
  completedAt: at(seconds),
});

test("durations are credited to the job that declares them", () => {
  const workflows = new Map<string, unknown>([
    [
      "caller.yml",
      {
        jobs: {
          // A matrix job named without its values gets them in parentheses.
          matrix: {
            name: "Matrix",
            strategy: { matrix: { shard: [1, 2] } },
            "timeout-minutes": 10,
            steps: [{ name: "Work", run: "true", "timeout-minutes": 2 }],
          },
          templated: {
            name: "Pack ${{ matrix.pkg }}",
            strategy: { matrix: { pkg: ["a", "b"] } },
            "timeout-minutes": 10,
          },
          call: { name: "Report", uses: "./callee.yml" },
          // Same displayed name twice: not credited to either.
          twin: { name: "Twin", "timeout-minutes": 5 },
          twin2: { name: "Twin", "timeout-minutes": 5 },
        },
      },
    ],
    [
      "callee.yml",
      { jobs: { file: { name: "File issues", "timeout-minutes": 5 } } },
    ],
  ]);
  const runs = [
    {
      workflow: "caller.yml",
      checkRuns: [
        checkRun("Matrix (1)", 100, "SUCCESS", [step("Work", 60)]),
        // A failed job is no duration, but its successful step still is.
        checkRun("Matrix (2)", 590, "FAILURE", [step("Work", 80)]),
        // A step that ran to its limit is no duration even when reported
        // successful, as continue-on-error reports it.
        checkRun("Matrix", 200, "SUCCESS", [step("Work", 120)]),
        checkRun("Pack a", 30),
        checkRun("Pack b", 50),
        checkRun("Pack c", 40, "SKIPPED"),
        checkRun("Report / File issues", 20),
        checkRun("Twin", 10),
        checkRun("cla", 3),
      ],
    },
  ];
  expect(aggregateWorkflowDurations(runs, workflows)).toEqual({
    "callee.yml": { file: { job: { p95Seconds: 20, samples: 1 } } },
    "caller.yml": {
      matrix: {
        job: { p95Seconds: 200, samples: 2 },
        steps: { Work: { p95Seconds: 80, samples: 2 } },
      },
      templated: { job: { p95Seconds: 50, samples: 2 } },
    },
  });
});

test("only completed runs are read, and a truncated page is refused", () => {
  const node = (status: string, total: number) => ({
    runs: {
      nodes: [
        {
          checkSuite: {
            status,
            checkRuns: {
              totalCount: total,
              nodes: [
                {
                  id: "CR_1",
                  name: "Job",
                  conclusion: "SUCCESS",
                  startedAt: at(0),
                  completedAt: at(5),
                },
              ],
            },
          },
        },
      ],
    },
  });
  expect(parseWorkflowRuns("w.yml", node("IN_PROGRESS", 1))).toEqual([]);
  expect(parseWorkflowRuns("w.yml", node("COMPLETED", 1))).toEqual([
    {
      workflow: "w.yml",
      checkRuns: [
        {
          id: "CR_1",
          name: "Job",
          conclusion: "SUCCESS",
          startedAt: at(0),
          completedAt: at(5),
          steps: [],
        },
      ],
    },
  ]);
  expect(() => parseWorkflowRuns("w.yml", node("COMPLETED", 2))).toThrow(
    "raise the page size",
  );
});

test("the timing refresh rewrites the measured durations", () => {
  const refresh = readFileSync(
    path.join(REPO_ROOT, ".github/workflows/test-timing-refresh.yml"),
    "utf8",
  );
  expect(refresh).toContain(
    `bun scripts/workflow-durations.ts \\\n            --repository "$GITHUB_REPOSITORY" \\\n            --output ${WORKFLOW_DURATIONS_FILE}`,
  );
  expect(refresh).toContain(`--workflow-durations ${WORKFLOW_DURATIONS_FILE}`);
});
