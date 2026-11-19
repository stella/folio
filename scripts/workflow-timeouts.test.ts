import { panic } from "better-result";
import { expect, setDefaultTimeout, test } from "bun:test";
import { existsSync, readFileSync } from "node:fs";
import path from "node:path";

import { scanTestTimeout } from "../test/scan-timeout";
import {
  aggregateWorkflowDurations,
  type CheckRun,
  committedWorkflows as workflowsByFile,
  fetchWorkflowRuns,
  limitsUnderMeasuredDurations,
  minimumMinutes,
  p95,
  parseWorkflowDurations,
  parseWorkflowRuns,
  readWorkflowDurations,
  WORKFLOW_DURATIONS_FILE,
  type WorkflowDurations,
} from "./workflow-durations";
import { derivedMinutes } from "./workflow-limit";

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

type CheckRunOptions = {
  name: string;
  seconds: number;
  conclusion?: string;
  steps?: CheckRun["steps"];
};
const checkRun = ({
  name,
  seconds,
  conclusion = "SUCCESS",
  steps = [],
}: CheckRunOptions): CheckRun => ({
  id: `${name}-${seconds}`,
  name,
  conclusion,
  startedAt: at(0),
  completedAt: at(seconds),
  steps,
});

type StepOptions = Pick<CheckRunOptions, "name" | "seconds" | "conclusion">;
const step = ({ name, seconds, conclusion = "SUCCESS" }: StepOptions) => ({
  name,
  conclusion,
  startedAt: at(0),
  completedAt: at(seconds),
});

const HISTORICAL_SHA = "c6d66dfa2a5ad1f96a21b819fe456fecdd3d6f6d";
const source = {
  repository: "stella/folio",
  sha: HISTORICAL_SHA,
  path: ".github/workflows/caller.yml",
};
const runFile = {
  repositoryName: source.repository,
  repositoryFileUrl: `https://github.com/${source.repository}/blob/${source.sha}/${source.path}`,
  path: source.path,
};

test("durations are credited to the job that declares them", () => {
  const workflows = new Map<string, unknown>([
    [
      ".github/workflows/caller.yml",
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
          call: { name: "Report", uses: "./.github/workflows/callee.yml" },
          // An unmatched declared name does not affect other measurements.
          twin: { name: "Twin", "timeout-minutes": 5 },
          twin2: { name: "Twin", "timeout-minutes": 5 },
        },
      },
    ],
    [
      ".github/workflows/callee.yml",
      { jobs: { file: { name: "File issues", "timeout-minutes": 5 } } },
    ],
  ]);
  const runs = [
    {
      source,
      workflows,
      checkRuns: [
        checkRun({
          name: "Matrix (1)",
          seconds: 100,
          conclusion: "SUCCESS",
          steps: [step({ name: "Work", seconds: 60 })],
        }),
        // A failed job is no duration, but its successful step still is.
        checkRun({
          name: "Matrix (2)",
          seconds: 590,
          conclusion: "FAILURE",
          steps: [step({ name: "Work", seconds: 80 })],
        }),
        // A step that ran to its limit is no duration even when reported
        // successful, as continue-on-error reports it.
        checkRun({
          name: "Matrix",
          seconds: 200,
          conclusion: "SUCCESS",
          steps: [step({ name: "Work", seconds: 120 })],
        }),
        checkRun({ name: "Pack a", seconds: 30 }),
        checkRun({ name: "Pack b", seconds: 50 }),
        checkRun({ name: "Pack c", seconds: 40, conclusion: "SKIPPED" }),
        checkRun({ name: "Report / File issues", seconds: 20 }),
        checkRun({ name: "cla", seconds: 3 }),
      ],
    },
  ];
  expect(aggregateWorkflowDurations(runs)).toEqual({
    ".github/workflows/callee.yml": {
      file: { job: { p95Seconds: 20, samples: 1 } },
    },
    ".github/workflows/caller.yml": {
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
          file: runFile,
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
  expect(parseWorkflowRuns(node("IN_PROGRESS", 1))).toEqual([]);
  expect(parseWorkflowRuns(node("COMPLETED", 1))).toEqual([
    {
      source,
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
  expect(() => parseWorkflowRuns(node("COMPLETED", 2))).toThrow(
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

const NIGHTLY_SETUP_REPORT_MINUTES = 7;

test("the nightly job covers every sequential phase cap plus setup and reporting", () => {
  const workflow = workflowsByFile().get(
    ".github/workflows/nightly-browser-input-fuzzer.yml",
  );
  const jobs = isRecord(workflow) ? workflow["jobs"] : undefined;
  const job = isRecord(jobs) ? jobs["browser-input-fuzzer"] : undefined;
  if (!isRecord(job) || !Array.isArray(job["steps"]))
    panic("The nightly browser job must declare its sequential steps.");
  const jobMinutes = job["timeout-minutes"];
  if (typeof jobMinutes !== "number")
    panic("The nightly browser job requires a numeric timeout.");
  const phases = job["steps"].flatMap((definition) => {
    if (
      !isRecord(definition) ||
      typeof definition["timeout-minutes"] !== "number"
    )
      return [];
    const name = definition["name"];
    if (typeof name !== "string")
      panic("Every timed nightly phase must have a name.");
    return [{ name, minutes: definition["timeout-minutes"] }];
  });
  expect(phases.length).toBeGreaterThan(0);
  const phaseMinutes = phases.reduce((sum, phase) => sum + phase.minutes, 0);
  expect(
    jobMinutes,
    phases.map(({ name, minutes }) => `${name}: ${minutes} min`).join("; "),
  ).toBeGreaterThanOrEqual(phaseMinutes + NIGHTLY_SETUP_REPORT_MINUTES);
});

test("timed names cannot select a duplicate sibling, including untimed ones", () => {
  for (const sibling of [{ "timeout-minutes": 2 }, {}]) {
    const workflows = new Map([
      [
        source.path,
        {
          jobs: {
            run: {
              "timeout-minutes": 10,
              steps: [
                { name: "Work", "timeout-minutes": 2 },
                { name: "Work", ...sibling },
              ],
            },
          },
        },
      ],
    ]);
    expect(limitsUnderMeasuredDurations(workflows, {})).toEqual([
      `${source.path}: run repeats timed step name "Work"`,
    ]);
    expect(() =>
      aggregateWorkflowDurations([
        {
          source,
          workflows,
          checkRuns: [
            checkRun({
              name: "run",
              seconds: 100,
              steps: [step({ name: "Work", seconds: 1 })],
            }),
          ],
        },
      ]),
    ).toThrow("Historical workflow limits are ambiguous");
  }
});

test("completed runs require an immutable matching executed-file reference", () => {
  for (const file of [
    null,
    {
      ...runFile,
      repositoryFileUrl: runFile.repositoryFileUrl.replace(
        HISTORICAL_SHA,
        "main",
      ),
    },
    { ...runFile, path: ".github/workflows/other.yml" },
    { ...runFile, repositoryName: "other/repository" },
    { ...runFile, path: "../caller.yml" },
  ])
    expect(() =>
      parseWorkflowRuns({
        runs: {
          nodes: [
            {
              file,
              checkSuite: {
                status: "COMPLETED",
                checkRuns: { nodes: [] },
              },
            },
          ],
        },
      }),
    ).toThrow("immutable executed workflow");
});

type FixtureTransportOptions = {
  files: ReadonlyMap<string, unknown>;
  runFiles?: readonly (typeof runFile)[];
};

/** GitHub-shaped transport: every workflow and check step crosses its parser. */
const fixtureTransport = ({
  files,
  runFiles = [runFile, runFile],
}: FixtureTransportOptions) => {
  const requestedFiles: string[] = [];
  const check = checkRun({ name: "Caller / Job", seconds: 125 });
  const fetchJson: Parameters<
    typeof fetchWorkflowRuns
  >[0]["fetchJson"] = async (url, body) => {
    if (body === undefined) {
      expect(url).toContain("/actions/workflows?");
      return {
        total_count: 1,
        workflows: [{ path: source.path, node_id: "W_1" }],
      };
    }
    const request: unknown = JSON.parse(body);
    if (
      !isRecord(request) ||
      typeof request["query"] !== "string" ||
      !isRecord(request["variables"])
    )
      panic("Malformed fixture request");
    const query = request["query"];
    const variables = request["variables"];
    if (query.includes("... on Workflow"))
      return {
        data: {
          w0: {
            runs: {
              nodes: runFiles.map((file) => ({
                file,
                checkSuite: {
                  status: "COMPLETED",
                  checkRuns: { totalCount: 1, nodes: [check] },
                },
              })),
            },
          },
        },
      };
    if (query.includes("... on Blob")) {
      const entries = Object.entries(variables).filter(([name]) =>
        name.startsWith("f"),
      );
      return {
        data: {
          repository: Object.fromEntries(
            entries.map(([alias, expression]) => {
              if (typeof expression !== "string")
                panic("Missing file expression");
              requestedFiles.push(expression);
              const definition = files.get(expression);
              return [
                alias,
                definition === undefined
                  ? null
                  : { text: JSON.stringify(definition) },
              ];
            }),
          ),
        },
      };
    }
    expect(query).toContain("... on CheckRun");
    return {
      data: {
        nodes: [
          {
            id: check.id,
            steps: {
              totalCount: 1,
              nodes: [step({ name: "Work", seconds: 125 })],
            },
          },
        ],
      },
    };
  };
  return { fetchJson, requestedFiles };
};

const CALLER = {
  jobs: { call: { name: "Caller", uses: "./.github/workflows/callee.yml" } },
};
const callee = (minutes: number) => ({
  jobs: {
    run: {
      name: "Job",
      "timeout-minutes": minutes,
      steps: [{ name: "Work", "timeout-minutes": minutes }],
    },
  },
});

test("historical caller and callee limits censor samples even after current limits rise", async () => {
  const nextSha = "8a823eda2b7f60a0b55f59e88f406a70f7a749fa";
  const nextFile = {
    ...runFile,
    repositoryFileUrl: runFile.repositoryFileUrl.replace(
      HISTORICAL_SHA,
      nextSha,
    ),
  };
  const oldCallee = ".github/workflows/callee.yml";
  const transport = fixtureTransport({
    runFiles: [runFile, runFile, nextFile],
    files: new Map([
      [`${HISTORICAL_SHA}:${source.path}`, CALLER],
      [`${HISTORICAL_SHA}:${oldCallee}`, callee(2)],
      [`${nextSha}:${source.path}`, CALLER],
      [`${nextSha}:${oldCallee}`, callee(10)],
    ]),
  });
  const current = new Map(
    Object.entries({
      [source.path]: CALLER,
      [oldCallee]: callee(30),
    }),
  );
  const runs = await fetchWorkflowRuns({
    repository: source.repository,
    workflows: current,
    fetchJson: transport.fetchJson,
  });
  // 125s exceeded the old two-minute cap; only the newer run is eligible.
  // A newly raised current cap must never turn the two censored runs healthy.
  expect(aggregateWorkflowDurations(runs)).toEqual({
    [oldCallee]: {
      run: {
        job: { p95Seconds: 125, samples: 1 },
        steps: { Work: { p95Seconds: 125, samples: 1 } },
      },
    },
  });
  expect(transport.requestedFiles.toSorted()).toEqual(
    [
      `${HISTORICAL_SHA}:${source.path}`,
      `${HISTORICAL_SHA}:${oldCallee}`,
      `${nextSha}:${source.path}`,
      `${nextSha}:${oldCallee}`,
    ].toSorted(),
  );
});

test("missing historical definitions and recursive caller cycles fail refresh explicitly", async () => {
  for (const files of [
    {},
    { [`${HISTORICAL_SHA}:${source.path}`]: CALLER },
    {
      [`${HISTORICAL_SHA}:${source.path}`]: CALLER,
      [`${HISTORICAL_SHA}:.github/workflows/callee.yml`]: {
        jobs: { cycle: { uses: `./${source.path}` } },
      },
    },
  ]) {
    const transport = fixtureTransport({
      files: new Map(Object.entries(files)),
    });
    await expect(
      fetchWorkflowRuns({
        repository: source.repository,
        workflows: new Map([[source.path, CALLER]]),
        fetchJson: transport.fetchJson,
      }),
    ).rejects.toThrow(/Missing historical workflow|cycle/u);
  }
});

test("missing requested check-run steps fail refresh rather than reduce the sample", async () => {
  for (const nodes of [[], [null], [{ id: "Caller / Job-125" }]]) {
    const transport = fixtureTransport({
      files: new Map([
        [`${HISTORICAL_SHA}:${source.path}`, CALLER],
        [`${HISTORICAL_SHA}:.github/workflows/callee.yml`, callee(10)],
      ]),
    });
    await expect(
      fetchWorkflowRuns({
        repository: source.repository,
        workflows: new Map([[source.path, CALLER]]),
        fetchJson: async (url, body) => {
          const response = await transport.fetchJson(url, body);
          return body?.includes("... on CheckRun")
            ? { data: { nodes } }
            : response;
        },
      }),
    ).rejects.toThrow(/requested check run|steps nodes/u);
  }
});

test("ambiguous timed job names fail attribution instead of losing samples", () => {
  const workflows = new Map([
    [
      source.path,
      {
        jobs: {
          first: { name: "Twin", "timeout-minutes": 5 },
          second: { name: "Twin", "timeout-minutes": 10 },
        },
      },
    ],
  ]);
  expect(() =>
    aggregateWorkflowDurations([
      {
        source,
        workflows,
        checkRuns: [checkRun({ name: "Twin", seconds: 100 })],
      },
    ]),
  ).toThrow("Ambiguous workflow job name");
});

test("duration snapshots require finite nonnegative percentiles and positive integer sample counts", () => {
  for (const measurement of [
    { p95Seconds: -1, samples: 1 },
    JSON.parse('{"p95Seconds":1e999,"samples":1}'),
    { p95Seconds: 10, samples: 0 },
    { p95Seconds: 10, samples: -1 },
    { p95Seconds: 10, samples: 0.5 },
  ])
    for (const durations of [
      { job: measurement },
      { steps: { Work: measurement } },
    ])
      expect(() =>
        parseWorkflowDurations({ [source.path]: { run: durations } }),
      ).toThrow("malformed");
  const valid = {
    [source.path]: { run: { job: { p95Seconds: 0, samples: 1 } } },
  };
  expect(parseWorkflowDurations(valid)).toEqual(valid);
});

const derivedWorkflow = (command: string, expression: string) => ({
  jobs: {
    run: {
      "timeout-minutes": 30,
      steps: [
        { id: "limit", name: "Derive", run: command },
        { name: "Timed", run: "true", "timeout-minutes": expression },
      ],
    },
  },
});
const LIMIT_COMMAND = `bun scripts/workflow-limit.ts --workflow fixture --job run --step "Timed" --fallback 5 --max 20 >> "$GITHUB_OUTPUT"`;
const LIMIT_EXPRESSION = "${{ fromJSON(steps.limit.outputs.minutes) }}";

test("a step limit derived from its own measurement is accepted and measured", () => {
  const workflow = derivedWorkflow(LIMIT_COMMAND, LIMIT_EXPRESSION);
  expect(
    limitsUnderMeasuredDurations(new Map([["fixture", workflow]]), {
      fixture: { run: { steps: { Timed: { p95Seconds: 9999, samples: 3 } } } },
    }),
  ).toEqual([]);
});

test("a derived limit must come from a step that measures the same step", () => {
  const other = LIMIT_COMMAND.replace('--step "Timed"', '--step "Other"');
  expect(
    limitsUnderMeasuredDurations(
      new Map([["fixture", derivedWorkflow(other, LIMIT_EXPRESSION)]]),
      {},
    ),
  ).toEqual([
    'fixture: run step "Timed" derives its limit from fixture run "Other"',
  ]);
  expect(
    limitsUnderMeasuredDurations(
      new Map([["fixture", derivedWorkflow("true", LIMIT_EXPRESSION)]]),
      {},
    ),
  ).toHaveLength(1);
});

test("the derived minutes follow the measurement and never fall under the fallback", () => {
  const target = { workflow: "fixture", job: "run", step: "Timed" };
  const measured = (p95Seconds: number): WorkflowDurations => ({
    fixture: { run: { steps: { Timed: { p95Seconds, samples: 3 } } } },
  });
  expect(derivedMinutes({}, target, 5, 20)).toBe(5);
  expect(derivedMinutes(measured(60), target, 5, 20)).toBe(5);
  expect(derivedMinutes(measured(600), target, 5, 20)).toBe(
    minimumMinutes(600),
  );
});

test("a measurement over the reviewed max fails with the way to raise it", () => {
  const target = { workflow: "fixture", job: "run", step: "Timed" };
  const slow: WorkflowDurations = {
    fixture: { run: { steps: { Timed: { p95Seconds: 1200, samples: 3 } } } },
  };
  expect(() => derivedMinutes(slow, target, 5, 20)).toThrow(
    "measured p95 20.0 min needs a 25 min limit, which exceeds the reviewed max 20; raise it in fixture if intended",
  );
  expect(derivedMinutes(slow, target, 5, 25)).toBe(25);
});
