// Binds the CI plan's areas to the jobs they gate. The planner (the shared
// stella/.github ci-plan action) selects areas from `.github/ci-plan.json`;
// `ci.yml` maps each area to a `ci-plan` output and gates a job on it. An area
// no job reads, an output no area feeds, or a job that runs without asking the
// plan would each drift in silence, so all three are checked here.
import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const PLAN_JOB = "ci-plan";
// The changeset gate runs beside the plan: it has its own pull-request-only
// condition and nothing to scope.
const UNPLANNED_JOBS = new Set([PLAN_JOB, "changeset", "ci-result"]);
const AREA_OUTPUT =
  /^\$\{\{ github\.event_name == 'merge_group' \|\| fromJSON\(steps\.plan\.outputs\.areas\)\.([a-z][a-z0-9_]*) \}\}$/u;
const GATE =
  /^needs\.ci-plan\.outputs\.([a-z][a-z0-9_]*_required) == '(true|false)'(?: && needs\.ci-plan\.outputs\.suite_depth == 'full')?$/u;

type Job = { needs?: unknown; if?: unknown; outputs?: Record<string, unknown> };

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const readWorkflow = (file: string) => {
  const workflow: unknown = Bun.YAML.parse(
    readFileSync(path.join(REPO_ROOT, ".github/workflows", file), "utf8"),
  );
  if (!isRecord(workflow)) throw new Error(`${file} is not a mapping`);
  return workflow;
};

const readJobs = (): Record<string, Job> => {
  const workflow = readWorkflow("ci.yml");
  if (!isRecord(workflow["jobs"])) throw new Error("ci.yml has no jobs");
  const jobs: Record<string, Job> = {};
  for (const [id, job] of Object.entries(workflow["jobs"])) {
    if (!isRecord(job)) throw new Error(`ci.yml job ${id} is not a mapping`);
    const outputs = job["outputs"];
    jobs[id] = { needs: job["needs"], if: job["if"], outputs: isRecord(outputs) ? outputs : {} };
  }
  return jobs;
};

const readPolicy = () => {
  const policy: unknown = JSON.parse(
    readFileSync(path.join(REPO_ROOT, ".github/ci-plan.json"), "utf8"),
  );
  if (!isRecord(policy) || !isRecord(policy["areas"])) throw new Error("ci-plan.json has no areas");
  return { areas: Object.keys(policy["areas"]), fullDepth: policy["fullDepth"] };
};

describe("CI plan", () => {
  const jobs = readJobs();
  const { areas, fullDepth } = readPolicy();
  const planOutputs = jobs[PLAN_JOB]?.outputs ?? {};

  // The shared planner's full-depth output can depend on labels or non-PR
  // events. Neither may promote this workflow's depth or expand its path scopes.
  test("only merge groups select full depth and bypass path scopes", () => {
    expect(planOutputs["suite_depth"]).toBe(
      "${{ github.event_name == 'merge_group' && 'full' || 'fast' }}",
    );
    expect(fullDepth).toBe("scoped");
  });

  // CodSpeed rejects merge_group events, so benchmarks keep their own
  // pull-request, main-push and nightly triggers.
  test("benchmarks never run on merge groups", () => {
    const triggers = readWorkflow("benchmarks.yml")["on"];
    if (!isRecord(triggers)) throw new Error("benchmarks.yml is missing triggers");
    expect(Object.keys(triggers)).not.toContain("merge_group");
    expect(Object.keys(triggers)).toContain("push");
  });

  for (const file of ["oracle-mutation-check.yml", "vscode-extension.yml"]) {
    test(`${file} runs heavy jobs only for merge groups`, () => {
      const workflow = readWorkflow(file);
      const triggers = workflow["on"];
      const heavyJobs = workflow["jobs"];
      if (!isRecord(triggers) || !isRecord(heavyJobs)) {
        throw new Error(`${file} is missing triggers or jobs`);
      }
      expect(Object.keys(triggers).toSorted()).toEqual(["merge_group", "pull_request"]);
      expect(triggers["pull_request"]).toEqual({
        types: ["opened", "synchronize", "reopened", "ready_for_review"],
      });
      expect(triggers["merge_group"]).toEqual({ types: ["checks_requested"] });
      expect(Object.keys(heavyJobs).length).toBeGreaterThan(0);
      for (const job of Object.values(heavyJobs)) {
        if (!isRecord(job)) throw new Error(`${file} contains an invalid job`);
        expect(job["if"]).toBe("github.event_name == 'merge_group'");
      }
    });
  }

  test("every area is a plan output named after it, and every area output is an area", () => {
    const mapped = Object.entries(planOutputs).flatMap(([output, value]) => {
      const area = typeof value === "string" ? AREA_OUTPUT.exec(value)?.[1] : undefined;
      return area === undefined ? [] : [{ output, area }];
    });
    expect(mapped.map(({ area }) => area).toSorted()).toEqual(areas.toSorted());
    for (const { output, area } of mapped) expect(output).toBe(`${area}_required`);
  });

  test("every other job waits on the plan and is gated by one area it maps", () => {
    const gated = new Set<string>();
    for (const [id, job] of Object.entries(jobs)) {
      if (UNPLANNED_JOBS.has(id)) continue;
      expect([id, job.needs]).toEqual([id, PLAN_JOB]);
      const gate = typeof job.if === "string" ? GATE.exec(job.if) : null;
      expect([id, gate !== null]).toEqual([id, true]);
      const output = gate?.[1] ?? "";
      expect([id, output in planOutputs]).toEqual([id, true]);
      if (gate?.[2] === "true") gated.add(output);
    }
    // No area is planned for nothing.
    expect([...gated].toSorted()).toEqual(areas.map((area) => `${area}_required`).toSorted());
  });
});
