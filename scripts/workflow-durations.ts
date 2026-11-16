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

type Workflows = ReadonlyMap<string, unknown>;

type WorkflowFile = { repository: string; sha: string; path: string };
type ParsedRun = { source: WorkflowFile; checkRuns: readonly CheckRun[] };

/** Each run is bound to the exact workflow definitions GitHub executed. */
export type MeasuredRun = ParsedRun & { workflows: Workflows };

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
    const names = new Map<string, number>();
    for (const step of steps) {
      const name = isRecord(step) ? literalName(step["name"]) : undefined;
      if (name !== undefined) names.set(name, (names.get(name) ?? 0) + 1);
    }
    const reported = new Set<string>();
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
      if (names.get(name) !== 1) {
        if (!reported.has(name))
          problems.push(`${file}: ${id} repeats timed step name "${name}"`);
        reported.add(name);
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
type JobNameSourcesOptions = {
  file: string;
  workflows: Workflows;
  seen?: ReadonlySet<string>;
};

const jobNameSources = ({
  file,
  workflows,
  seen = new Set(),
}: JobNameSourcesOptions): { source: string; target: JobTarget }[] =>
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
    const visited = new Set(seen);
    visited.add(file);
    return jobNameSources({ file: callee, workflows, seen: visited }).map(
      (inner) => ({
        source: `${source} / ${inner.source}`,
        target: inner.target,
      }),
    );
  });

/** Resolves a check run name of `file`'s runs to the job that declared it. */
export const jobForCheckRun = (file: string, workflows: Workflows) => {
  const patterns = jobNameSources({ file, workflows }).map(
    ({ source, target }) => ({
      pattern: new RegExp(`^${source}$`, "u"),
      target,
    }),
  );
  return (name: string): JobTarget | undefined => {
    const matches = patterns.filter(({ pattern }) => pattern.test(name));
    if (matches.length > 1)
      return panic(`Ambiguous workflow job name "${name}" in ${file}.`);
    return matches.at(0)?.target;
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
  for (const [file, workflow] of workflows) {
    const { limits, problems } = declaredLimits(file, workflow);
    if (problems.length > 0)
      panic(`Historical workflow limits are ambiguous: ${problems.join("; ")}`);
    for (const { job, minutes, step } of limits) {
      const key = keyOf({ workflow: file, job });
      byJob.set(key, (byJob.get(key) ?? new Map()).set(step, minutes));
    }
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
): WorkflowDurations => {
  const samples = new Map<string, Map<string | undefined, number[]>>();
  type AddSampleOptions = {
    key: string;
    step?: string;
    seconds: number;
    minutes: number | undefined;
  };
  const add = ({ key, step, seconds, minutes }: AddSampleOptions) => {
    // No historical numeric limit means no eligible capped measurement.
    // Never treat missing history as an unlimited healthy sample.
    if (minutes === undefined || seconds >= minutes * 60) return;
    const byStep = samples.get(key) ?? new Map<string | undefined, number[]>();
    const values = byStep.get(step) ?? [];
    values.push(seconds);
    byStep.set(step, values);
    samples.set(key, byStep);
  };
  const resolvers = new Map<string, (name: string) => JobTarget | undefined>();
  for (const { source, checkRuns, workflows } of runs) {
    const workflow = source.path;
    const sourceId = sourceKey(source);
    const limits = limitsByJob(workflows);
    const resolve =
      resolvers.get(sourceId) ?? jobForCheckRun(workflow, workflows);
    resolvers.set(sourceId, resolve);
    for (const checkRun of checkRuns) {
      const target = resolve(checkRun.name);
      if (target === undefined) continue;
      const key = keyOf(target);
      const seconds = secondsOf(checkRun);
      if (seconds !== undefined)
        add({ key, seconds, minutes: limits.get(key)?.get(undefined) });
      for (const name of limits.get(key)?.keys() ?? []) {
        if (name === undefined) continue;
        const step = checkRun.steps.find(
          (candidate) => candidate.name === name,
        );
        const stepSeconds = step === undefined ? undefined : secondsOf(step);
        if (stepSeconds !== undefined)
          add({
            key,
            step: name,
            seconds: stepSeconds,
            minutes: limits.get(key)?.get(name),
          });
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
    const jobs = result[file] ?? {};
    jobs[job] = durations;
    result[file] = jobs;
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
  Number.isFinite(value["p95Seconds"]) &&
  value["p95Seconds"] >= 0 &&
  typeof value["samples"] === "number" &&
  Number.isInteger(value["samples"]) &&
  value["samples"] > 0;

export const parseWorkflowDurations = (parsed: unknown): WorkflowDurations => {
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

export const readWorkflowDurations = (root = REPO_ROOT): WorkflowDurations =>
  parseWorkflowDurations(
    JSON.parse(readFileSync(path.join(root, WORKFLOW_DURATIONS_FILE), "utf8")),
  );

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

const WORKFLOW_PATH = /^\.github\/workflows\/[^/]+\.ya?ml$/u;
const WORKFLOW_BLOB =
  /^https:\/\/github\.com\/([^/]+\/[^/]+)\/blob\/([a-f0-9]{40})\/(.+)$/u;

const parseWorkflowFile = (value: unknown): WorkflowFile => {
  const workflowPath = field(value, "path");
  const repository = field(value, "repositoryName");
  const url = field(value, "repositoryFileUrl");
  const match = typeof url === "string" ? WORKFLOW_BLOB.exec(url) : null;
  const sha = match?.at(2);
  if (
    typeof workflowPath !== "string" ||
    !WORKFLOW_PATH.test(workflowPath) ||
    typeof repository !== "string" ||
    match?.at(1) !== repository ||
    match?.at(3) !== workflowPath ||
    sha === undefined
  )
    return panic(
      "A completed workflow run must identify its immutable executed workflow file.",
    );
  return { repository, sha, path: workflowPath };
};

const sourceKey = ({ repository, sha, path: workflowPath }: WorkflowFile) =>
  `${repository}\n${sha}\n${workflowPath}`;
const revisionKey = ({ repository, sha }: WorkflowFile) =>
  `${repository}\n${sha}`;

/** Completed runs retain GitHub's executed-file reference for historical limits. */
export const parseWorkflowRuns = (node: unknown): ParsedRun[] =>
  nodesOf(field(node, "runs"), "runs").flatMap((run) => {
    const suite = field(run, "checkSuite");
    if (field(suite, "status") !== "COMPLETED") return [];
    return [
      {
        source: parseWorkflowFile(field(run, "file")),
        checkRuns: nodesOf(field(suite, "checkRuns"), "check runs").map(
          (checkRun) => {
            const id = field(checkRun, "id");
            if (typeof id !== "string")
              return panic("GitHub check run has no id");
            return Object.assign({ id, steps: [] }, parseStep(checkRun));
          },
        ),
      },
    ];
  });

type WorkflowIdsOptions = {
  repository: string;
  workflows: Workflows;
  fetchJson: FetchJson;
};

const workflowIds = async ({
  repository,
  workflows,
  fetchJson,
}: WorkflowIdsOptions) => {
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

type GraphqlOptions = {
  fetchJson: FetchJson;
  query: string;
  variables: Record<string, unknown>;
};

const graphql = async ({ fetchJson, query, variables }: GraphqlOptions) => {
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

type HistoricalWorkflowsOptions = {
  sources: readonly WorkflowFile[];
  fetchJson: FetchJson;
};

const historicalWorkflows = async ({
  sources,
  fetchJson,
}: HistoricalWorkflowsOptions) => {
  const definitions = new Map<string, Map<string, unknown>>();
  const pending = new Map(sources.map((source) => [sourceKey(source), source]));
  while (pending.size > 0) {
    const requested = [...pending.values()];
    pending.clear();
    const repositories = new Map<string, WorkflowFile[]>();
    for (const source of requested) {
      if (definitions.get(revisionKey(source))?.has(source.path)) continue;
      const files = repositories.get(source.repository) ?? [];
      files.push(source);
      repositories.set(source.repository, files);
    }
    for (const [repository, files] of repositories) {
      const [owner, name, extra] = repository.split("/");
      if (!owner || !name || extra !== undefined)
        return panic("Workflow repository must have an owner and name.");
      for (const batch of chunks(files, CHECK_RUN_BATCH)) {
        const data = await graphql({
          fetchJson,
          query: `query($owner: String!, $name: String!, ${batch.map((_, index) => `$f${index}: String!`).join(", ")}) { repository(owner: $owner, name: $name) { ${batch.map((_, index) => `f${index}: object(expression: $f${index}) { ... on Blob { text } }`).join(" ")} } }`,
          variables: {
            owner,
            name,
            ...Object.fromEntries(
              batch.map((source, index) => [
                `f${index}`,
                `${source.sha}:${source.path}`,
              ]),
            ),
          },
        });
        for (const [index, source] of batch.entries()) {
          const text = field(
            field(field(data, "repository"), `f${index}`),
            "text",
          );
          if (typeof text !== "string")
            return panic(
              `Missing historical workflow ${source.repository}@${source.sha}:${source.path}.`,
            );
          const workflow: unknown = Bun.YAML.parse(text);
          if (!isRecord(workflow) || !isRecord(workflow["jobs"]))
            return panic(
              `Historical workflow ${source.path} has no jobs mapping.`,
            );
          const key = revisionKey(source);
          const snapshot = definitions.get(key) ?? new Map<string, unknown>();
          snapshot.set(source.path, workflow);
          definitions.set(key, snapshot);
          for (const { job } of jobsOf(workflow)) {
            const callee = localCallee(job);
            if (callee === undefined) continue;
            if (!WORKFLOW_PATH.test(callee))
              return panic(
                `Invalid historical reusable workflow path ${callee}.`,
              );
            const next = {
              repository: source.repository,
              sha: source.sha,
              path: callee,
            };
            if (!snapshot.has(callee)) pending.set(sourceKey(next), next);
          }
        }
      }
    }
  }
  // A visited definition is fetched once; a caller cycle is invalid provenance.
  const check = (source: WorkflowFile, visiting: ReadonlySet<string>): void => {
    if (visiting.has(source.path))
      return panic(`Historical reusable workflow cycle at ${source.path}.`);
    const snapshot = definitions.get(revisionKey(source));
    if (snapshot === undefined || !snapshot.has(source.path))
      return panic("Every historical workflow must have a loaded definition.");
    const next = new Set(visiting);
    next.add(source.path);
    for (const { job } of jobsOf(snapshot.get(source.path))) {
      const callee = localCallee(job);
      if (callee !== undefined)
        check(
          { repository: source.repository, sha: source.sha, path: callee },
          next,
        );
    }
  };
  for (const source of sources) check(source, new Set());
  return definitions;
};

type FetchWorkflowRunsOptions = WorkflowIdsOptions;
export const fetchWorkflowRuns = async ({
  repository,
  workflows,
  fetchJson,
}: FetchWorkflowRunsOptions): Promise<MeasuredRun[]> => {
  const ids = await workflowIds({ repository, workflows, fetchJson });
  const parsed: ParsedRun[] = [];
  for (const batch of chunks([...ids.keys()].toSorted(), WORKFLOW_BATCH)) {
    const data = await graphql({
      fetchJson,
      query: `query(${batch.map((_, index) => `$w${index}: ID!`).join(", ")}) { ${batch
        .map(
          (_, index) =>
            `w${index}: node(id: $w${index}) { ... on Workflow { runs(first: ${RUNS_PER_WORKFLOW}, orderBy: {field: CREATED_AT, direction: DESC}) { nodes { file { path repositoryName repositoryFileUrl } checkSuite { status checkRuns(first: 100) { totalCount nodes { id ${STEP_FIELDS} } } } } } } }`,
        )
        .join(" ")} }`,
      variables: Object.fromEntries(
        batch.map((file, index) => [`w${index}`, ids.get(file)]),
      ),
    });
    for (const [index] of batch.entries())
      parsed.push(...parseWorkflowRuns(field(data, `w${index}`)));
  }
  const snapshots = await historicalWorkflows({
    sources: parsed.map(({ source }) => source),
    fetchJson,
  });
  const runs: MeasuredRun[] = parsed.map(({ source, checkRuns }) => ({
    source,
    checkRuns,
    workflows:
      snapshots.get(revisionKey(source)) ??
      panic("A parsed run must have historical workflow definitions."),
  }));
  const wanted = [
    ...new Set(
      runs.flatMap(({ source, checkRuns, workflows: snapshot }) => {
        const stepNames = measuredStepNames(snapshot);
        const resolve = jobForCheckRun(source.path, snapshot);
        return checkRuns.flatMap((checkRun) => {
          const target = resolve(checkRun.name);
          return target !== undefined &&
            stepNames.has(keyOf(target)) &&
            checkRun.conclusion !== "SKIPPED" &&
            checkRun.completedAt !== null
            ? [checkRun.id]
            : [];
        });
      }),
    ),
  ];
  const steps = new Map<string, CheckStep[]>();
  for (const batch of chunks(wanted, CHECK_RUN_BATCH)) {
    const data = await graphql({
      fetchJson,
      query: `query($ids: [ID!]!) { nodes(ids: $ids) { ... on CheckRun { id steps(first: 100) { totalCount nodes { ${STEP_FIELDS} } } } } }`,
      variables: { ids: batch },
    });
    const nodes = field(data, "nodes");
    if (!Array.isArray(nodes)) panic("GitHub returned no check run nodes");
    for (const node of nodes) {
      const id = field(node, "id");
      if (typeof id !== "string" || !batch.includes(id) || steps.has(id))
        return panic("GitHub returned an invalid requested check run node.");
      steps.set(
        id,
        nodesOf(field(node, "steps"), "steps").map((step) => parseStep(step)),
      );
    }
    for (const id of batch)
      if (!steps.has(id))
        return panic(`GitHub omitted requested check run ${id}.`);
  }
  return runs.map(({ source, workflows: snapshot, checkRuns }) => ({
    source,
    workflows: snapshot,
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
  const runs = await fetchWorkflowRuns({
    repository,
    workflows,
    fetchJson: githubFetchJson(valueFor("--cache")),
  });
  writeFileSync(
    output,
    `${JSON.stringify(aggregateWorkflowDurations(runs), null, 2)}\n`,
  );
}
