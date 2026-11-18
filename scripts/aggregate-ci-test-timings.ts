#!/usr/bin/env bun
import { panic } from "better-result";
import { readFileSync, writeFileSync } from "node:fs";
import path from "node:path";

import { plannedTestFiles, type TestTimings } from "./ci-test-shard-plan";
import {
  fullSuiteTimingSets,
  readFileTimings,
} from "./ci-test-timing-artifact";
import { propertyFiles } from "./property-areas";
import { WORKFLOW_DURATIONS_FILE } from "./workflow-durations";

type TimingArtifact = { runId: string; timings: TestTimings };

type TimingArtifactReference = { artifactId: number; runId: number };

type FetchJson = (url: string) => Promise<unknown>;

type WorkflowRun = { event: string; id: number };

type Artifact = { id: number; name: string; expired?: boolean };

/** Where a timing file's measurements come from. */
export type TimingSource = {
  /** The committed file the aggregate replaces. */
  output: string;
  workflow: string;
  /** Run-listing filters for that workflow. */
  runs: string;
  events: readonly string[];
  /** Artifacts named `<prefix>-<run id>-<shard>`. */
  prefix: string;
  /** Unmeasured files get the median (scheduling) or stay absent (planned from budgets). */
  fillUnmeasured: boolean;
  files: () => string[];
};

const REPO_ROOT = path.resolve(import.meta.dir, "..");

export const TIMING_SOURCES = {
  // Candidates from successful full runs; each envelope proves its own shard set.
  tests: {
    output: "scripts/ci-test-timings.json",
    workflow: "ci.yml",
    runs: "status=success",
    events: ["merge_group", "push"],
    prefix: "test-timings",
    fillUnmeasured: true,
    files: () => plannedTestFiles(REPO_ROOT),
  },
  // Extended property runs on main. A property failure does not change how
  // long its file ran, and a shard cut off at its timeout uploads nothing, so
  // any completed run with a shard's measurements counts.
  properties: {
    output: "scripts/property-test-timings.json",
    workflow: "property-proofs.yml",
    runs: "status=completed&branch=main",
    events: ["workflow_dispatch"],
    prefix: "property-timings",
    fillUnmeasured: false,
    files: () => propertyFiles().map(({ file }) => file),
  },
} as const satisfies Record<string, TimingSource>;

export type TimingSourceName = keyof typeof TIMING_SOURCES;

const isTimingSourceName = (name: string): name is TimingSourceName =>
  Object.hasOwn(TIMING_SOURCES, name);

type GraphqlRequestOptions = {
  base: string;
  branch: string;
  repository: string;
  /** Committed repository path to the local file holding its contents. */
  timingFiles: Record<string, string>;
};

export const createGraphqlRequest = ({
  base,
  branch,
  repository,
  timingFiles,
}: GraphqlRequestOptions) => ({
  query:
    "mutation($input: CreateCommitOnBranchInput!) { createCommitOnBranch(input: $input) { commit { oid } } }",
  variables: {
    input: {
      branch: { repositoryNameWithOwner: repository, branchName: branch },
      expectedHeadOid: base,
      message: { headline: "perf(ci): refresh measured test timings" },
      fileChanges: {
        additions: Object.entries(timingFiles).map(([repoPath, file]) => ({
          path: repoPath,
          contents: readFileSync(file).toString("base64"),
        })),
      },
    },
  },
});

const recordsFor = (value: unknown, key: string) => {
  if (typeof value !== "object" || value === null || !(key in value)) {
    panic(`GitHub response is missing ${key}`);
  }
  const records = value[key as keyof typeof value];
  if (!Array.isArray(records))
    panic(`GitHub response field ${key} is not an array`);
  return records;
};

export const selectTimingArtifacts = async (
  repository: string,
  fetchJson: FetchJson,
  source: TimingSource = TIMING_SOURCES.tests,
): Promise<TimingArtifactReference[]> => {
  const selected: TimingArtifactReference[] = [];
  let accepted = 0;
  // Timing artifacts expire after 14 days, so ten pages of runs reach past every usable set.
  for (let page = 1; page <= 10 && accepted < 5; page++) {
    const runsResponse = await fetchJson(
      `https://api.github.com/repos/${repository}/actions/workflows/${source.workflow}/runs?${source.runs}&per_page=100&page=${page}`,
    );
    const pageRuns = recordsFor(runsResponse, "workflow_runs");
    const runs = pageRuns.filter(
      (run): run is WorkflowRun =>
        typeof run === "object" &&
        run !== null &&
        typeof run.id === "number" &&
        source.events.includes(run.event),
    );
    for (const run of runs) {
      const artifactsResponse = await fetchJson(
        `https://api.github.com/repos/${repository}/actions/runs/${run.id}/artifacts?per_page=100`,
      );
      const prefix = `${source.prefix}-${run.id}-`;
      const artifacts = recordsFor(artifactsResponse, "artifacts").filter(
        (artifact): artifact is Artifact =>
          typeof artifact === "object" &&
          artifact !== null &&
          typeof artifact.id === "number" &&
          typeof artifact.name === "string" &&
          artifact.name.startsWith(prefix) &&
          artifact.expired !== true,
      );
      if (artifacts.length === 0) continue;
      selected.push(
        ...artifacts.map(({ id }) => ({ artifactId: id, runId: run.id })),
      );
      accepted += 1;
      if (accepted === 5) break;
    }
    if (pageRuns.length < 100) break;
  }
  return selected;
};

const githubFetchJson: FetchJson = async (url) => {
  const token = process.env.GH_TOKEN;
  if (!token) panic("GH_TOKEN is required to list timing artifacts");
  const response = await fetch(url, {
    headers: {
      Accept: "application/vnd.github+json",
      Authorization: `Bearer ${token}`,
      "X-GitHub-Api-Version": "2022-11-28",
    },
  });
  if (!response.ok)
    panic(`GitHub request failed with ${response.status}: ${url}`);
  return response.json();
};

const median = (values: number[]) => {
  const sorted = values.toSorted((a, b) => a - b);
  const middle = Math.floor(sorted.length / 2);
  if (sorted.length === 0) panic("Cannot aggregate empty timing samples");
  if (sorted.length % 2 === 1) return sorted.at(middle) ?? 0;
  return Math.round(
    ((sorted.at(middle - 1) ?? 0) + (sorted.at(middle) ?? 0)) / 2,
  );
};

export const aggregateTestTimings = (
  artifacts: TimingArtifact[],
  files: string[],
  { fillUnmeasured = true }: { fillUnmeasured?: boolean } = {},
): string => {
  const runs = new Map<string, Map<string, number>>();
  for (const { runId, timings } of artifacts) {
    const run = runs.get(runId) ?? new Map<string, number>();
    for (const [file, duration] of Object.entries(timings))
      run.set(file, duration);
    runs.set(runId, run);
  }
  const allMeasurements = [...runs.values()].flatMap((run) =>
    Array.from(run.values()),
  );
  const fallback = fillUnmeasured ? median(allMeasurements) : undefined;
  const result: TestTimings = {};
  for (const file of [...new Set(files)].toSorted()) {
    const samples = [...runs.values()].flatMap((run) => {
      const duration = run.get(file);
      return duration === undefined ? [] : [duration];
    });
    if (samples.length > 0) result[file] = median(samples);
    else if (fallback !== undefined) result[file] = fallback;
  }
  return `${JSON.stringify(result, null, 2)}\n`;
};

const readPropertyTimingArtifact = (value: unknown): TimingArtifact => {
  if (
    typeof value !== "object" ||
    value === null ||
    !("runId" in value) ||
    typeof value.runId !== "string" ||
    !("timings" in value)
  )
    return panic(
      "Property timing artifact requires a run id and file measurements",
    );
  return {
    runId: value.runId,
    timings: readFileTimings(value.timings, value.runId),
  };
};

// The `--graphql-request` flag naming each source's local file.
const COMMIT_FLAGS: Record<TimingSourceName, string> = {
  tests: "--timings",
  properties: "--property-timings",
};

if (import.meta.main) {
  const args = process.argv.slice(2);
  const valueFor = (name: string) => {
    const index = args.indexOf(name);
    return index === -1 ? undefined : args.at(index + 1);
  };
  const sourceName = valueFor("--source") ?? "tests";
  if (!isTimingSourceName(sourceName))
    panic(`Unknown timing source ${sourceName}`);
  const source = TIMING_SOURCES[sourceName];
  const artifactManifest = valueFor("--artifact-manifest");
  if (artifactManifest !== undefined) {
    const repository = valueFor("--repository");
    if (!artifactManifest || !repository) {
      panic(
        "Usage: bun scripts/aggregate-ci-test-timings.ts --artifact-manifest <file> --repository <owner/name> [--source tests|properties]",
      );
    }
    const artifacts = await selectTimingArtifacts(
      repository,
      githubFetchJson,
      source,
    );
    writeFileSync(
      artifactManifest,
      artifacts
        .map(({ artifactId, runId }) => `${runId}\t${artifactId}`)
        .join("\n") + (artifacts.length === 0 ? "" : "\n"),
    );
    process.exit(0);
  }
  const requestOutput = valueFor("--graphql-request");
  if (requestOutput !== undefined) {
    const timingFiles: Record<string, string> = {};
    for (const [name, flag] of Object.entries(COMMIT_FLAGS)) {
      const file = valueFor(flag);
      if (file !== undefined && isTimingSourceName(name))
        timingFiles[TIMING_SOURCES[name].output] = file;
    }
    const durations = valueFor("--workflow-durations");
    if (durations !== undefined)
      timingFiles[WORKFLOW_DURATIONS_FILE] = durations;
    const repository = valueFor("--repository");
    const branch = valueFor("--branch");
    const base = valueFor("--base");
    if (
      !requestOutput ||
      Object.keys(timingFiles).length === 0 ||
      !repository ||
      !branch ||
      !base
    ) {
      panic(
        "Usage: bun scripts/aggregate-ci-test-timings.ts --graphql-request <file> [--timings <file>] [--property-timings <file>] [--workflow-durations <file>] --repository <owner/name> --branch <branch> --base <oid>",
      );
    }
    writeFileSync(
      requestOutput,
      `${JSON.stringify(createGraphqlRequest({ base, branch, repository, timingFiles }))}\n`,
    );
    process.exit(0);
  }
  const output = valueFor("--output");
  const valueFlags = new Set(["--output", "--source", "--artifact-root"]);
  const inputs = args.filter(
    (arg, index) =>
      !valueFlags.has(arg) &&
      !(index > 0 && valueFlags.has(args[index - 1] ?? "")),
  );
  if (!output || inputs.length === 0) {
    panic(
      "Usage: bun scripts/aggregate-ci-test-timings.ts --output <file> [--source tests|properties] [--artifact-root <directory>] <artifacts...>",
    );
  }
  const inputsRead = inputs.map((input): unknown =>
    JSON.parse(readFileSync(input, "utf8")),
  );
  const artifacts = (() => {
    if (sourceName === "properties")
      return inputsRead.map(readPropertyTimingArtifact);
    const root = valueFor("--artifact-root");
    if (!root)
      return panic(
        "Full-suite timing inputs require their download manifest --artifact-root",
      );
    const proven = fullSuiteTimingSets(
      inputsRead.map((artifact, index) => {
        const input =
          inputs.at(index) ?? panic("A timing artifact lost its input path");
        const relative = path.relative(path.resolve(root), path.resolve(input));
        if (
          relative.startsWith(`..${path.sep}`) ||
          relative === ".." ||
          path.isAbsolute(relative)
        )
          return panic("A timing input is outside its download manifest root", {
            input,
          });
        const runId =
          relative.split(path.sep).at(0) ??
          panic("A timing input lost its manifest run id");
        return { runId, artifact };
      }),
    );
    if (proven.legacyRuns.length > 0)
      console.warn(
        `Skipped timing runs without planned shard provenance: ${proven.legacyRuns.join(", ")}`,
      );
    if (proven.artifacts.length === 0) {
      console.warn(
        `No proven full-suite timing sets are available; keeping ${output}.`,
      );
      process.exit(0);
    }
    return proven.artifacts;
  })();
  writeFileSync(
    output,
    aggregateTestTimings(artifacts, source.files(), {
      fillUnmeasured: source.fillUnmeasured,
    }),
  );
}
