#!/usr/bin/env bun
import { panic } from "better-result";
import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";

// Measured durations of every job and every step that declares its own
// timeout-minutes, from recent runs of each workflow. The timeout rule in
// scripts/workflow-timeouts.test.ts reads the committed file, and the test
// timing refresh (.github/workflows/test-timing-refresh.yml) rewrites it.

const REPO_ROOT = path.resolve(import.meta.dir, "..");

export const WORKFLOW_DURATIONS_FILE = "scripts/workflow-durations.json";

/** A healthy run's 95th percentile uses at most this share of its limit. */
export const TIMEOUT_SHARE = 0.8;

/** Recent runs per workflow that the refresh measures. */
const RUNS_PER_WORKFLOW = 50;

export type Measured = { p95Seconds: number; samples: number };

/** A job's own duration and those of its steps that declare a limit. */
export type JobDurations = { job?: Measured; steps?: Record<string, Measured> };

/** Workflow file, then job id, as committed in WORKFLOW_DURATIONS_FILE. */
export type WorkflowDurations = Record<string, Record<string, JobDurations>>;

/** A check run or one of its steps, as GitHub reports it. */
export type CheckStep = {
  name: string;
  conclusion: string | null;
  startedAt: string | null;
  completedAt: string | null;
};

/** One check run, a job's run; its steps are fetched only where measured. */
export type CheckRun = CheckStep & { id: string; steps: readonly CheckStep[] };

/** The check runs of one completed workflow run. */
export type MeasuredRun = { workflow: string; checkRuns: readonly CheckRun[] };

type Workflows = ReadonlyMap<string, unknown>;

type JobTarget = { workflow: string; job: string };

type Limit = JobTarget & { minutes: number; step?: string };

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const field = (value: unknown, key: string): unknown =>
  isRecord(value) ? value[key] : undefined;

const jobsOf = (workflow: unknown) =>
  isRecord(workflow) && isRecord(workflow["jobs"])
    ? Object.entries(workflow["jobs"]).flatMap(([id, job]) =>
        isRecord(job) ? [{ id, job }] : [],
      )
    : [];

const LOCAL_CALLEE = /^\.\/(\.github\/workflows\/[^@]+\.ya?ml)$/u;

const localCallee = (job: Record<string, unknown>) => {
  const uses = job["uses"];
  return typeof uses === "string" ? LOCAL_CALLEE.exec(uses)?.[1] : undefined;
};

/** A step's name when it is literal, so GitHub reports it unchanged. */
const literalName = (value: unknown) =>
  typeof value === "string" && !value.includes("${{") ? value : undefined;

/** The smallest whole-minute limit whose share covers a measured p95. */
export const minimumMinutes = (p95Seconds: number) =>
  Math.ceil(p95Seconds / (60 * TIMEOUT_SHARE));

type DeclaredLimits = { limits: Limit[]; problems: string[] };

const declaredLimits = (file: string, workflow: unknown): DeclaredLimits => {
  const limits: Limit[] = [];
  const problems: string[] = [];
  for (const { id, job } of jobsOf(workflow)) {
    const minutes = job["timeout-minutes"];
    if (typeof minutes === "number")
      limits.push({ workflow: file, job: id, minutes });
    const steps = Array.isArray(job["steps"]) ? job["steps"] : [];
    for (const [index, step] of steps.entries()) {
      if (!isRecord(step) || typeof step["timeout-minutes"] !== "number")
        continue;
      const name = literalName(step["name"]);
      if (name === undefined) {
        problems.push(
          `${file}: ${id} step ${index + 1} declares timeout-minutes without a literal name to measure it by`,
        );
        continue;
      }
      limits.push({
        workflow: file,
        job: id,
        minutes: step["timeout-minutes"],
        step: name,
      });
    }
  }
  return { limits, problems };
};

const escapeSource = (text: string) =>
  text.replaceAll(/[.*+?^${}()|[\]\\]/gu, String.raw`\$&`);

// GitHub names a job's check run by its `name` (expressions evaluated) or its
// id, adds the matrix values in parentheses unless the name uses them, and
// names a reusable workflow's jobs `<caller> / <callee job>`.
const jobNameSources = (
  file: string,
  workflows: Workflows,
  seen: ReadonlySet<string> = new Set(),
): { source: string; target: JobTarget }[] =>
  jobsOf(workflows.get(file)).flatMap(({ id, job }) => {
    const label = typeof job["name"] === "string" ? job["name"] : id;
    const usesExpression = label.includes("${{");
    const template = label
      .split(/\$\{\{.*?\}\}/u)
      .map(escapeSource)
      .join(".+");
    const matrix =
      isRecord(job["strategy"]) &&
      job["strategy"]["matrix"] !== undefined &&
      !usesExpression;
    const source = matrix ? `${template}(?: \\(.+\\))?` : template;
    const callee = localCallee(job);
    if (callee === undefined || seen.has(callee))
      return [{ source, target: { workflow: file, job: id } }];
    return jobNameSources(callee, workflows, new Set([...seen, file])).map(
      (inner) => ({
        source: `${source} / ${inner.source}`,
        target: inner.target,
      }),
    );
  });

/** Resolves a check run name of `file`'s runs to the job that declared it. */
export const jobForCheckRun = (file: string, workflows: Workflows) => {
  const patterns = jobNameSources(file, workflows).map(
    ({ source, target }) => ({
      pattern: new RegExp(`^${source}$`, "u"),
      target,
    }),
  );
  return (name: string): JobTarget | undefined => {
    const matches = patterns.filter(({ pattern }) => pattern.test(name));
    // An ambiguous name is not measured rather than credited to a guess.
    return matches.length === 1 ? matches[0]?.target : undefined;
  };
};

const secondsOf = ({ conclusion, startedAt, completedAt }: CheckStep) => {
  if (conclusion !== "SUCCESS" || startedAt === null || completedAt === null)
    return undefined;
  const seconds = (Date.parse(completedAt) - Date.parse(startedAt)) / 1000;
  return Number.isFinite(seconds) && seconds >= 0 ? seconds : undefined;
};

/** Nearest-rank 95th percentile. */
export const p95 = (values: readonly number[]) => {
  const sorted = values.toSorted((a, b) => a - b);
  const value = sorted.at(Math.ceil(sorted.length * 0.95) - 1);
  if (value === undefined) panic("Cannot take a percentile of no samples");
  return value;
};

const measured = (samples: readonly number[]): Measured => ({
  p95Seconds: p95(samples),
  samples: samples.length,
});

const keyOf = ({ workflow, job }: JobTarget) => `${workflow}\n${job}`;

/** Each job's declared minutes and its steps' (by name), keyed by keyOf. */
const limitsByJob = (workflows: Workflows) => {
  const byJob = new Map<string, Map<string | undefined, number>>();
  for (const [file, workflow] of workflows)
    for (const { job, minutes, step } of declaredLimits(file, workflow)
      .limits) {
      const key = keyOf({ workflow: file, job });
      byJob.set(key, (byJob.get(key) ?? new Map()).set(step, minutes));
    }
  return byJob;
};

/** Step names with their own limit, by the job (keyOf) that declares them. */
const measuredStepNames = (workflows: Workflows) =>
  new Map(
    [...limitsByJob(workflows)].flatMap(([key, limits]) => {
      const names = [...limits.keys()].filter((name) => name !== undefined);
      return names.length === 0 ? [] : [[key, names] as const];
    }),
  );

// Only a successful job or step is a duration: a skipped or failed one is
// not, and neither is one that ran to its limit, which a step with
// continue-on-error still reports as successful. A step counts even when a
// later step failed its job.
export const aggregateWorkflowDurations = (
  runs: readonly MeasuredRun[],
  workflows: Workflows,
): WorkflowDurations => {
  const limits = limitsByJob(workflows);
  const samples = new Map<string, Map<string | undefined, number[]>>();
  const add = (key: string, step: string | undefined, seconds: number) => {
    const minutes = limits.get(key)?.get(step);
    if (minutes !== undefined && seconds >= minutes * 60) return;
    const byStep = samples.get(key) ?? new Map<string | undefined, number[]>();
    byStep.set(step, [...(byStep.get(step) ?? []), seconds]);
    samples.set(key, byStep);
  };
  const resolvers = new Map<string, (name: string) => JobTarget | undefined>();
  for (const { workflow, checkRuns } of runs) {
    const resolve =
      resolvers.get(workflow) ?? jobForCheckRun(workflow, workflows);
    resolvers.set(workflow, resolve);
    for (const checkRun of checkRuns) {
      const target = resolve(checkRun.name);
      if (target === undefined) continue;
      const key = keyOf(target);
      const seconds = secondsOf(checkRun);
      if (seconds !== undefined) add(key, undefined, seconds);
      for (const name of limits.get(key)?.keys() ?? []) {
        if (name === undefined) continue;
        const step = checkRun.steps.find(
          (candidate) => candidate.name === name,
        );
        const stepSeconds = step === undefined ? undefined : secondsOf(step);
        if (stepSeconds !== undefined) add(key, name, stepSeconds);
      }
    }
  }
  const result: WorkflowDurations = {};
  for (const key of [...samples.keys()].toSorted()) {
    const [file = "", job = ""] = key.split("\n");
    const byStep = samples.get(key) ?? new Map<string | undefined, number[]>();
    const own = byStep.get(undefined);
    const steps = [...byStep.entries()]
      .flatMap(([name, values]) =>
        name === undefined ? [] : [[name, measured(values)] as const],
      )
      .toSorted(([a], [b]) => a.localeCompare(b));
    const durations: JobDurations = {};
    if (own !== undefined) durations.job = measured(own);
    if (steps.length > 0) durations.steps = Object.fromEntries(steps);
    result[file] = { ...result[file], [job]: durations };
  }
  return result;
};

/**
 * Every job or step limit a healthy run can reach: its declared minutes are
 * fewer than its measured p95 divided by TIMEOUT_SHARE.
 */
export const limitsUnderMeasuredDurations = (
  workflows: Workflows,
  durations: WorkflowDurations,
): string[] =>
  [...workflows].flatMap(([file, workflow]) => {
    const { limits, problems } = declaredLimits(file, workflow);
    return problems.concat(
      limits.flatMap(({ job, minutes, step }) => {
        const jobDurations = durations[file]?.[job];
        const measurement =
          step === undefined ? jobDurations?.job : jobDurations?.steps?.[step];
        if (measurement === undefined) return [];
        const needed = minimumMinutes(measurement.p95Seconds);
        return minutes >= needed
          ? []
          : [
              `${file}: ${step === undefined ? job : `${job} step "${step}"`} allows ${minutes} min, but its measured p95 of ${measurement.p95Seconds} s needs at least ${needed} min`,
            ];
      }),
    );
  });

export const committedWorkflows = (root = REPO_ROOT): Map<string, unknown> =>
  new Map(
    [...new Bun.Glob(".github/workflows/*.{yml,yaml}").scanSync({ cwd: root })]
      .toSorted()
      .map((file) => [
        file,
        Bun.YAML.parse(readFileSync(path.join(root, file), "utf8")),
      ]),
  );

const isMeasured = (value: unknown): value is Measured =>
  isRecord(value) &&
  typeof value["p95Seconds"] === "number" &&
  typeof value["samples"] === "number";

export const readWorkflowDurations = (root = REPO_ROOT): WorkflowDurations => {
  const parsed: unknown = JSON.parse(
    readFileSync(path.join(root, WORKFLOW_DURATIONS_FILE), "utf8"),
  );
  if (!isRecord(parsed)) panic(`${WORKFLOW_DURATIONS_FILE} is not an object`);
  const result: WorkflowDurations = {};
  for (const [file, jobs] of Object.entries(parsed)) {
    if (!isRecord(jobs)) panic(`${file} durations are not an object`);
    const entries: Record<string, JobDurations> = {};
    for (const [job, value] of Object.entries(jobs)) {
      const own = field(value, "job");
      const steps = field(value, "steps");
      if (own !== undefined && !isMeasured(own))
        panic(`${file} ${job} has a malformed job duration`);
      if (steps !== undefined && !isRecord(steps))
        panic(`${file} ${job} has malformed step durations`);
      const entry: JobDurations = {};
      if (own !== undefined) entry.job = own;
      if (steps !== undefined)
        entry.steps = Object.fromEntries(
          Object.entries(steps).map(([name, step]) => {
            if (!isMeasured(step))
              panic(`${file} ${job} step ${name} has a malformed duration`);
            return [name, step];
          }),
        );
      entries[job] = entry;
    }
    result[file] = entries;
  }
  return result;
};

// ---------------------------------------------------------------------------
// Refresh: GraphQL requests sent one at a time. Job durations come from each
// workflow's recent runs; steps only for the check runs of jobs that declare a
// step limit, fetched by id, since steps for every check run are slow to serve.

type FetchJson = (url: string, body?: string) => Promise<unknown>;

const STEP_FIELDS = "name conclusion startedAt completedAt";

/** Workflows per request for job durations. */
const WORKFLOW_BATCH = 5;

/** Check runs per request for step durations. */
const CHECK_RUN_BATCH = 25;

const nodesOf = (value: unknown, label: string): unknown[] => {
  const nodes = field(value, "nodes");
  const total = field(value, "totalCount");
  if (!Array.isArray(nodes)) panic(`GitHub response has no ${label} nodes`);
  if (typeof total === "number" && total > nodes.length)
    panic(
      `GitHub returned ${nodes.length} of ${total} ${label}; raise the page size`,
    );
  return nodes;
};

const nullableString = (value: unknown) =>
  typeof value === "string" ? value : null;

const parseStep = (value: unknown): CheckStep => {
  const name = field(value, "name");
  if (typeof name !== "string") panic("GitHub check run has no name");
  return {
    name,
    conclusion: nullableString(field(value, "conclusion")),
    startedAt: nullableString(field(value, "startedAt")),
    completedAt: nullableString(field(value, "completedAt")),
  };
};

/** The completed runs of one workflow node in a GraphQL response. */
export const parseWorkflowRuns = (
  workflow: string,
  node: unknown,
): MeasuredRun[] =>
  nodesOf(field(node, "runs"), "runs").flatMap((run) => {
    const suite = field(run, "checkSuite");
    if (field(suite, "status") !== "COMPLETED") return [];
    return [
      {
        workflow,
        checkRuns: nodesOf(field(suite, "checkRuns"), "check runs").map(
          (checkRun) => {
            const id = field(checkRun, "id");
            if (typeof id !== "string") panic("GitHub check run has no id");
            const { name, conclusion, startedAt, completedAt } =
              parseStep(checkRun);
            return { id, name, conclusion, startedAt, completedAt, steps: [] };
          },
        ),
      },
    ];
  });

const workflowIds = async (
  repository: string,
  workflows: Workflows,
  fetchJson: FetchJson,
) => {
  const listing = await fetchJson(
    `https://api.github.com/repos/${repository}/actions/workflows?per_page=100`,
  );
  const listed = field(listing, "workflows");
  const total = field(listing, "total_count");
  if (!Array.isArray(listed) || (typeof total === "number" && total > 100))
    panic("GitHub workflow listing is missing or exceeds one page");
  const ids = new Map<string, string>();
  for (const entry of listed) {
    const file = field(entry, "path");
    const id = field(entry, "node_id");
    if (
      typeof file === "string" &&
      typeof id === "string" &&
      workflows.has(file)
    )
      ids.set(file, id);
  }
  return ids;
};

const graphql = async (
  fetchJson: FetchJson,
  query: string,
  variables: Record<string, unknown>,
) => {
  const response = await fetchJson(
    "https://api.github.com/graphql",
    JSON.stringify({ query, variables }),
  );
  const errors = field(response, "errors");
  if (errors !== undefined)
    panic(`GitHub GraphQL request failed: ${JSON.stringify(errors)}`);
  return field(response, "data");
};

const chunks = <T>(items: readonly T[], size: number) =>
  Array.from({ length: Math.ceil(items.length / size) }, (_, index) =>
    items.slice(index * size, (index + 1) * size),
  );

export const fetchWorkflowRuns = async (
  repository: string,
  workflows: Workflows,
  fetchJson: FetchJson,
): Promise<MeasuredRun[]> => {
  const ids = await workflowIds(repository, workflows, fetchJson);
  const runs: MeasuredRun[] = [];
  for (const batch of chunks([...ids.keys()].toSorted(), WORKFLOW_BATCH)) {
    const data = await graphql(
      fetchJson,
      `query(${batch.map((_, index) => `$w${index}: ID!`).join(", ")}) { ${batch
        .map(
          (_, index) =>
            `w${index}: node(id: $w${index}) { ... on Workflow { runs(first: ${RUNS_PER_WORKFLOW}, orderBy: {field: CREATED_AT, direction: DESC}) { nodes { checkSuite { status checkRuns(first: 100) { totalCount nodes { id ${STEP_FIELDS} } } } } } } }`,
        )
        .join(" ")} }`,
      Object.fromEntries(
        batch.map((file, index) => [`w${index}`, ids.get(file)]),
      ),
    );
    for (const [index, file] of batch.entries())
      runs.push(...parseWorkflowRuns(file, field(data, `w${index}`)));
  }
  const stepNames = measuredStepNames(workflows);
  const wanted = runs.flatMap(({ workflow, checkRuns }) => {
    const resolve = jobForCheckRun(workflow, workflows);
    return checkRuns.flatMap((checkRun) => {
      const target = resolve(checkRun.name);
      return target !== undefined &&
        stepNames.has(keyOf(target)) &&
        checkRun.conclusion !== "SKIPPED" &&
        checkRun.completedAt !== null
        ? [checkRun.id]
        : [];
    });
  });
  const steps = new Map<string, CheckStep[]>();
  for (const batch of chunks(wanted, CHECK_RUN_BATCH)) {
    const data = await graphql(
      fetchJson,
      `query($ids: [ID!]!) { nodes(ids: $ids) { ... on CheckRun { id steps(first: 100) { totalCount nodes { ${STEP_FIELDS} } } } } }`,
      { ids: batch },
    );
    const nodes = field(data, "nodes");
    if (!Array.isArray(nodes)) panic("GitHub returned no check run nodes");
    for (const node of nodes) {
      const id = field(node, "id");
      if (typeof id === "string")
        steps.set(
          id,
          nodesOf(field(node, "steps"), "steps").map((step) => parseStep(step)),
        );
    }
  }
  return runs.map(({ workflow, checkRuns }) => ({
    workflow,
    checkRuns: checkRuns.map(
      ({ id, name, conclusion, startedAt, completedAt }) => ({
        id,
        name,
        conclusion,
        startedAt,
        completedAt,
        steps: steps.get(id) ?? [],
      }),
    ),
  }));
};

const githubFetchJson =
  (cache: string | undefined): FetchJson =>
  async (url, body) => {
    const cached =
      cache === undefined
        ? undefined
        : path.join(
            cache,
            `${Bun.hash(`${url}\n${body ?? ""}`).toString(16)}.json`,
          );
    if (cached !== undefined && existsSync(cached))
      return JSON.parse(readFileSync(cached, "utf8"));
    const token = process.env["GH_TOKEN"];
    if (!token) panic("GH_TOKEN is required to measure workflow durations");
    const send = () =>
      fetch(url, {
        method: body === undefined ? "GET" : "POST",
        body,
        headers: {
          Accept: "application/vnd.github+json",
          Authorization: `Bearer ${token}`,
          "X-GitHub-Api-Version": "2022-11-28",
        },
      });
    // One pass at a steady pace; a gateway timeout on a large page is retried once.
    await Bun.sleep(1000);
    let response = await send();
    if (response.status >= 502 && response.status <= 504) {
      await Bun.sleep(10_000);
      response = await send();
    }
    if (!response.ok)
      panic(`GitHub request failed with ${response.status}: ${url}`);
    const json: unknown = await response.json();
    if (cached !== undefined) {
      mkdirSync(path.dirname(cached), { recursive: true });
      writeFileSync(cached, JSON.stringify(json));
    }
    return json;
  };

if (import.meta.main) {
  const args = process.argv.slice(2);
  const valueFor = (name: string) => {
    const index = args.indexOf(name);
    return index === -1 ? undefined : args.at(index + 1);
  };
  const repository = valueFor("--repository");
  const output = valueFor("--output");
  if (!repository || !output)
    panic(
      "Usage: bun scripts/workflow-durations.ts --repository <owner/name> --output <file> [--cache <directory>]",
    );
  const workflows = committedWorkflows();
  const runs = await fetchWorkflowRuns(
    repository,
    workflows,
    githubFetchJson(valueFor("--cache")),
  );
  writeFileSync(
    output,
    `${JSON.stringify(aggregateWorkflowDurations(runs, workflows), null, 2)}\n`,
  );
}
