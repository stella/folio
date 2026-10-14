import { expect, test } from "bun:test";
import { readdirSync, readFileSync } from "node:fs";
import path from "node:path";

const ROOT = path.resolve(import.meta.dir, "..");
const WORKFLOWS = path.join(ROOT, ".github/workflows");
const FAILURE_COMMAND =
  /scripts\/consumer-scenarios\.ts|canonical-session\.fuzz\.ts|--project=(?:parity|browser|interleaving)-fuzzer/u;
const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);
const record = (value: unknown, context: string): Record<string, unknown> => {
  if (!isRecord(value)) throw new TypeError(`${context} must be a mapping`);
  return value;
};
const parseWorkflow = (source: string, context: string) => {
  const parsed: unknown = Bun.YAML.parse(source);
  const workflow = record(parsed, context);
  return {
    name: workflow["name"],
    events: record(workflow["on"], `${context}.on`),
    nonEmitter: source.match(/^# fuzz-report: none -- (.+)$/mu)?.at(1),
    jobs: record(workflow["jobs"], `${context}.jobs`),
  };
};

type Upload = { name: string; paths: string[]; always: boolean };
type Reporter = {
  needs: string[];
  artifactPattern: string;
  logPattern: string;
  condition: string;
  permissions: Record<string, unknown>;
};

const listOf = (value: unknown): string[] => {
  if (typeof value === "string")
    return value
      .split("\n")
      .map((item) => item.trim())
      .filter(Boolean);
  if (Array.isArray(value)) return value.filter((item): item is string => typeof item === "string");
  return [];
};

const uploadsFor = (job: Record<string, unknown>): Upload[] => {
  if (!Array.isArray(job["steps"])) return [];
  return job["steps"].flatMap((stepValue) => {
    if (!isRecord(stepValue) || typeof stepValue["uses"] !== "string") return [];
    if (!stepValue["uses"].startsWith("actions/upload-artifact@")) return [];
    const config = isRecord(stepValue["with"]) ? stepValue["with"] : {};
    if (typeof config["name"] !== "string") return [];
    return [
      {
        name: config["name"],
        paths: listOf(config["path"]),
        always: typeof stepValue["if"] === "string" && stepValue["if"].includes("always()"),
      },
    ];
  });
};

const isFailureProducer = (step: Record<string, unknown>): boolean => {
  const run = typeof step["run"] === "string" ? step["run"] : "";
  return step["continue-on-error"] === true || FAILURE_COMMAND.test(run);
};

const logPathsFor = (job: Record<string, unknown>): string[] => {
  if (!Array.isArray(job["steps"])) return [];
  return job["steps"].flatMap((stepValue) => {
    if (!isRecord(stepValue) || typeof stepValue["run"] !== "string") return [];
    const env = isRecord(stepValue["env"]) ? stepValue["env"] : {};
    return [...stepValue["run"].matchAll(/\btee\s+(?:"([^"]+)"|'([^']+)'|([^\s|]+))/gu)].flatMap(
      (match) => {
        let logPath = match[1] ?? match[2] ?? match[3];
        if (!logPath) return [];
        for (const [key, value] of Object.entries(env)) {
          if (typeof value !== "string") continue;
          logPath = logPath.replaceAll(`$${key}`, value);
        }
        logPath = logPath.replaceAll("${{ github.workspace }}/", "");
        return logPath.includes("$") ? [] : [logPath];
      },
    );
  });
};

const reportersFor = (jobs: Record<string, unknown>): Reporter[] =>
  Object.values(jobs).flatMap((jobValue) => {
    if (!isRecord(jobValue) || jobValue["uses"] !== "./.github/workflows/fuzz-failure-report.yml")
      return [];
    const inputs = isRecord(jobValue["with"]) ? jobValue["with"] : {};
    if (typeof inputs["artifact-pattern"] !== "string") return [];
    const needs =
      typeof jobValue["needs"] === "string" ? [jobValue["needs"]] : listOf(jobValue["needs"]);
    return [
      {
        needs,
        artifactPattern: inputs["artifact-pattern"],
        logPattern: typeof inputs["log-pattern"] === "string" ? inputs["log-pattern"] : "**/*.log",
        condition: typeof jobValue["if"] === "string" ? jobValue["if"] : "",
        permissions: isRecord(jobValue["permissions"]) ? jobValue["permissions"] : {},
      },
    ];
  });

const workflowDefects = (jobs: Record<string, unknown>): string[] => {
  const reporters = reportersFor(jobs);
  const defects: string[] = [];
  for (const [jobId, jobValue] of Object.entries(jobs)) {
    if (!isRecord(jobValue) || !Array.isArray(jobValue["steps"])) continue;
    const producerSteps = jobValue["steps"].filter(
      (step): step is Record<string, unknown> => isRecord(step) && isFailureProducer(step),
    );
    if (producerSteps.length === 0) continue;
    const logs = logPathsFor(jobValue);
    const uploads = uploadsFor(jobValue);
    if (logs.length === 0) {
      defects.push(`${jobId}: producer has no captured log`);
      continue;
    }
    for (const log of logs) {
      const upload = uploads.find(
        ({ paths, always }) =>
          always &&
          paths.some(
            (artifactPath) =>
              log === artifactPath || log.startsWith(`${artifactPath.replace(/\/$/u, "")}/`),
          ),
      );
      if (!upload) {
        defects.push(`${jobId}: no uploaded artifact contains ${log}`);
        continue;
      }
      const matchingReporter = reporters.some((reporter) => {
        const artifactMatches = new Bun.Glob(reporter.artifactPattern).match(upload.name);
        const logMatches = new Bun.Glob(reporter.logPattern).match(log);
        return (
          reporter.needs.includes(jobId) &&
          reporter.condition.includes("github.ref == 'refs/heads/main'") &&
          reporter.permissions["issues"] === "write" &&
          artifactMatches &&
          logMatches
        );
      });
      if (!matchingReporter)
        defects.push(`${jobId}: no shared reporter covers ${upload.name}/${log}`);
    }
  }
  return defects;
};

const nightlyWorkflows = () =>
  readdirSync(WORKFLOWS)
    .filter((file) => file.endsWith(".yml"))
    .flatMap((file) => {
      const parsed = parseWorkflow(readFileSync(path.join(WORKFLOWS, file), "utf8"), file);
      if (typeof parsed.name !== "string" || !parsed.name.startsWith("Nightly")) return [];
      return [{ file, ...parsed, name: parsed.name }];
    });

type NightlyWorkflowOptions = { jobs: Record<string, unknown>; nonEmitter?: string };
const nightlyDefects = ({ jobs, nonEmitter }: NightlyWorkflowOptions): string[] => {
  if (nonEmitter !== undefined) {
    const producers = Object.values(jobs).filter(
      (job) =>
        isRecord(job) &&
        Array.isArray(job["steps"]) &&
        job["steps"].some((step) => isRecord(step) && isFailureProducer(step)),
    );
    return producers.length === 0 ? [] : ["non-emitter declaration contains a marker producer"];
  }
  const missing =
    reportersFor(jobs).length === 0
      ? ["nightly workflow needs a shared reporter or non-emitter declaration"]
      : [];
  return [...missing, ...workflowDefects(jobs)];
};

test("every nightly fuzz or flow producer has a scoped reusable failure reporter", () => {
  const workflows = nightlyWorkflows();
  expect(workflows.map(({ file }) => file).toSorted()).toContain("nightly-corpus-gate.yml");
  const defects = workflows.flatMap(({ file, jobs, nonEmitter }) =>
    nightlyDefects({ jobs, nonEmitter }).map((defect) => `${file}: ${defect}`),
  );
  expect(defects).toEqual([]);
  const webkit = parseWorkflow(
    readFileSync(path.join(WORKFLOWS, "nightly-webkit-parity.yml"), "utf8"),
    "webkit",
  );
  expect(Object.keys(webkit.events)).toEqual(["workflow_dispatch"]);
  expect(webkit.name).not.toMatch(/^Nightly/u);
});

test("the workflow guard rejects missing reporters and incorrect artifact log scopes", () => {
  const producer = {
    steps: [
      { run: "bun scripts/consumer-scenarios.ts 2>&1 | tee output/flow.log" },
      {
        if: "always()",
        uses: "actions/upload-artifact@v7",
        with: { name: "flow-evidence", path: "output" },
      },
    ],
  };
  const reporter = (logPattern: string) => ({
    needs: "flow",
    uses: "./.github/workflows/fuzz-failure-report.yml",
    if: "failure() && github.ref == 'refs/heads/main'",
    permissions: { contents: "read", issues: "write" },
    with: { "artifact-pattern": "flow-evidence", "log-pattern": logPattern },
  });
  expect(workflowDefects({ flow: producer })).toEqual([
    "flow: no shared reporter covers flow-evidence/output/flow.log",
  ]);
  expect(workflowDefects({ flow: producer, report: reporter("**/other/*.log") })).toEqual([
    "flow: no shared reporter covers flow-evidence/output/flow.log",
  ]);
  expect(workflowDefects({ flow: producer, report: reporter("**/*.log") })).toEqual([]);
});

test("shared reporter queues writes, preserves matrix artifacts, and requires failure markers", () => {
  const workflow = parseWorkflow(
    readFileSync(path.join(WORKFLOWS, "fuzz-failure-report.yml"), "utf8"),
    "reporter",
  );
  const job = record(workflow.jobs["report"], "report job");
  expect(job["permissions"]).toEqual({ contents: "read", issues: "write" });
  expect(job["concurrency"]).toEqual({
    group: "fuzz-failure-issues",
    queue: "max",
    "cancel-in-progress": false,
  });
  if (!Array.isArray(job["steps"])) throw new TypeError("Missing report steps");
  const steps = job["steps"].map((step) => record(step, "step"));
  const download = steps.find((step) =>
    String(step["uses"]).startsWith("actions/download-artifact@"),
  );
  expect(record(download?.["with"], "download inputs")["merge-multiple"]).toBe(false);
  expect(download?.["continue-on-error"]).toBeUndefined();
  const filing = steps.find((step) =>
    String(step["run"]).includes("scripts/fuzz-failure-issues.ts"),
  );
  expect(filing?.["run"]).toContain("--require-failures");
  expect(filing?.["run"]).toContain('--run-url "$RUN_URL"');
  expect(filing?.["run"]).toContain('--log-pattern "$LOG_PATTERN"');
});

test("unfamiliar nightly runners must declare reporting, and exemptions cannot hide marker producers", () => {
  expect(nightlyDefects({ jobs: { opaque: { steps: [{ run: "bun run new:fuzz" }] } } })).toEqual([
    "nightly workflow needs a shared reporter or non-emitter declaration",
  ]);
  expect(
    nightlyDefects({
      jobs: { marker: { steps: [{ run: "bun scripts/consumer-scenarios.ts" }] } },
      nonEmitter: "baseline only",
    }),
  ).toEqual(["non-emitter declaration contains a marker producer"]);
});
