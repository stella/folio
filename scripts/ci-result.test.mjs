import { describe, expect, test } from "bun:test";
import { mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

const workflow = Bun.YAML.parse(
  readFileSync(fileURLToPath(new URL("../.github/workflows/ci.yml", import.meta.url)), "utf8"),
);
const resultStep = workflow.jobs["ci-result"].steps[0];
const scopes = JSON.parse(resultStep.env.JOB_SCOPES);

const planOutputs = Object.keys(workflow.jobs["ci-plan"].outputs);
// Derived from the plan so a new selection area cannot be left unset here.
const areaOutputs = planOutputs.filter(
  (output) => output.endsWith("_required") && output !== "code_required",
);

const needsFor = ({ depth, event, codeRequired = "true", toolchainChanged = "false" }) => {
  const outputs = {
    ...Object.fromEntries(areaOutputs.map((area) => [area, "true"])),
    trusted: "true",
    suite_depth: depth,
    code_required: codeRequired,
    toolchain_changed: toolchainChanged,
    reused_run: "",
    warm_caches: "false",
  };
  const needs = { "ci-plan": { result: "success", outputs } };
  for (const [job, scope] of Object.entries(scopes)) {
    const selected =
      (scope.event === undefined || scope.event === event) &&
      scope.reuse !== "required" &&
      scope.warm !== "required" &&
      (scope.area === undefined ||
        (outputs[scope.area] === (scope.value ?? "true") &&
          (scope.depth === undefined || scope.depth === depth)));
    needs[job] = { result: selected ? "success" : "skipped" };
  }
  return needs;
};

const runResult = (event, needs, draft = "false") => {
  const directory = mkdtempSync(join(tmpdir(), "folio-ci-outcome-"));
  const summary = join(directory, "summary");
  try {
    const result = Bun.spawnSync(
      ["bash", "-c", `gh() { echo "$DRAFT_STATE"; }; ${resultStep.run}`],
      {
        env: {
          ...process.env,
          EVENT: event,
          NEEDS: JSON.stringify(needs),
          PLAN: JSON.stringify(needs["ci-plan"].outputs),
          PLAN_RESULT: needs["ci-plan"].result,
          JOB_SCOPES: JSON.stringify(scopes),
          DRAFT_STATE: draft,
          REPO: "stella/folio",
          PR_NUMBER: "1",
          GITHUB_STEP_SUMMARY: summary,
        },
      },
    );
    return { ...result, summary: Bun.file(summary).size > 0 ? readFileSync(summary, "utf8") : "" };
  } finally {
    rmSync(directory, { recursive: true, force: true });
  }
};

describe("CI result", () => {
  test("the aggregate observes every job and each scope matches the workflow gate", () => {
    const jobs = workflow.jobs;
    expect(
      Object.keys(jobs)
        .filter((job) => job !== "ci-plan" && job !== "ci-result")
        .toSorted(),
    ).toEqual(Object.keys(scopes).toSorted());
    expect(jobs["ci-result"].needs.toSorted()).toEqual(
      Object.keys(jobs)
        .filter((job) => job !== "ci-result")
        .toSorted(),
    );
    expect(workflow.jobs["ci-result"].steps).toHaveLength(1);
    for (const [job, scope] of Object.entries(scopes)) {
      const gate = jobs[job].if;
      if (scope.event) {
        expect(gate).toContain(`github.event_name == '${scope.event}'`);
        if (scope.event === "pull_request")
          expect(gate).toContain("github.event.pull_request.draft != true");
      }
      if (scope.area) {
        expect(planOutputs).toContain(scope.area);
        expect(gate).toContain(`needs.ci-plan.outputs.${scope.area} == '${scope.value ?? "true"}'`);
        expect(gate.includes("needs.ci-plan.outputs.suite_depth == 'full'")).toBe(
          scope.depth === "full",
        );
      }
    }
  });

  const scenarios = [
    ["fast", "pull_request", "true", "false"],
    ["fast", "pull_request", "false", "false"],
    ["fast", "push", "true", "false"],
    ["fast", "workflow_dispatch", "true", "false"],
    ["fast", "schedule", "true", "false"],
    ["full", "merge_group", "true", "false"],
    ["full", "pull_request", "true", "true"],
    ["full", "push", "true", "true"],
    ["full", "workflow_dispatch", "true", "true"],
    ["full", "schedule", "true", "true"],
  ];
  test("every declared job is selected in at least one scenario", () => {
    const exercised = new Set();
    for (const [depth, event, codeRequired, toolchainChanged] of scenarios) {
      for (const [job, { result }] of Object.entries(
        needsFor({ depth, event, codeRequired, toolchainChanged }),
      )) {
        if (job !== "ci-plan" && result === "success") exercised.add(job);
      }
    }
    for (const [job, { result }] of Object.entries(reusedNeeds("true"))) {
      if (job !== "ci-plan" && result === "success") exercised.add(job);
    }
    expect([...exercised].toSorted()).toEqual(Object.keys(scopes).toSorted());
  });

  for (const [depth, event, codeRequired, toolchainChanged] of scenarios) {
    const baseline = needsFor({ depth, event, codeRequired, toolchainChanged });
    const scenario = `${depth} ${event} code=${codeRequired} toolchain=${toolchainChanged}`;
    test(`successful selected jobs approve ${scenario}`, () => {
      expect(runResult(event, baseline).exitCode).toBe(0);
    });
    for (const [job, { result }] of Object.entries(baseline)) {
      if (job === "ci-plan" || result !== "success") continue;
      test(`${job} must succeed for ${scenario}`, () => {
        for (const failure of ["failure", "cancelled", "skipped", "missing"]) {
          const needs =
            failure === "missing"
              ? Object.fromEntries(Object.entries(baseline).filter(([name]) => name !== job))
              : structuredClone(baseline);
          if (failure !== "missing") needs[job].result = failure;
          const run = runResult(event, needs);
          expect(run.exitCode).toBe(1);
          expect(run.stdout.toString()).toContain(`${job} (selected, ${failure})`);
        }
      });
    }
  }

  const reusedNeeds = (warm = "false") => {
    const needs = needsFor({ depth: "fast", event: "push" });
    for (const key of planOutputs.filter((output) => output.endsWith("_required"))) {
      needs["ci-plan"].outputs[key] = "false";
    }
    needs["ci-plan"].outputs.reused_run = "38238759782";
    needs["ci-plan"].outputs.warm_caches = warm;
    for (const job of Object.keys(scopes)) needs[job].result = "skipped";
    if (warm === "true") needs["warm-main-caches"].result = "success";
    return needs;
  };

  test("a reused push accepts skipped correctness jobs and links its successful run", () => {
    for (const warm of ["false", "true"]) {
      const result = runResult("push", reusedNeeds(warm));
      expect(result.exitCode).toBe(0);
      expect(result.summary).toBe(
        "Reused merge-group run [38238759782](https://github.com/stella/folio/actions/runs/38238759782)\n",
      );
    }
  });

  test("all-skipped pushes require valid reuse and fulfilled cache warming", () => {
    for (const reused of ["", "0", "invalid", undefined]) {
      const needs = reusedNeeds();
      needs["ci-plan"].outputs.reused_run = reused;
      expect(runResult("push", needs).exitCode).toBe(1);
    }
    for (const event of ["pull_request", "merge_group", "workflow_dispatch"]) {
      expect(runResult(event, reusedNeeds()).exitCode).toBe(1);
    }
    for (const result of ["failure", "cancelled", "skipped", "missing"]) {
      const needs = reusedNeeds("true");
      needs["warm-main-caches"].result = result;
      expect(runResult("push", needs).exitCode).toBe(1);
    }
    for (const area of planOutputs.filter((key) => key.endsWith("_required"))) {
      const needs = reusedNeeds();
      needs["ci-plan"].outputs[area] = "true";
      expect(runResult("push", needs).exitCode).toBe(1);
    }
    for (const warm of ["", "invalid"]) {
      expect(runResult("push", reusedNeeds(warm)).exitCode).toBe(1);
    }
    const untrusted = reusedNeeds();
    untrusted["ci-plan"].outputs.trusted = "false";
    expect(runResult("push", untrusted).exitCode).toBe(1);
  });

  for (const job of Object.keys(scopes).filter((key) => key !== "warm-main-caches")) {
    test(`reused pushes require ${job} to be present and skipped`, () => {
      for (const result of ["success", "failure", "cancelled", "missing"]) {
        const baseline = reusedNeeds();
        const needs =
          result === "missing"
            ? Object.fromEntries(Object.entries(baseline).filter(([name]) => name !== job))
            : baseline;
        if (result !== "missing") needs[job].result = result;
        expect(runResult("push", needs).exitCode).toBe(1);
      }
    });
  }

  test("an unexpected failure, cancelled plan, or wrong event depth fails", () => {
    const needs = needsFor({ depth: "fast", event: "pull_request" });
    needs["docx-kernel"].result = "failure";
    expect(runResult("pull_request", needs).exitCode).toBe(1);
    needs["ci-plan"].result = "cancelled";
    expect(runResult("pull_request", needs).exitCode).toBe(1);

    for (const event of ["pull_request", "push", "workflow_dispatch", "schedule"]) {
      for (const depth of ["full", "invalid"]) {
        const run = runResult(event, needsFor({ depth, event }));
        expect(run.exitCode).toBe(1);
        expect(run.stdout.toString()).toContain(`wrong suite depth for ${event}`);
      }
    }
    for (const depth of ["fast", "invalid"]) {
      const run = runResult("merge_group", needsFor({ depth, event: "merge_group" }));
      expect(run.exitCode).toBe(1);
      expect(run.stdout.toString()).toContain("wrong suite depth for merge_group");
    }

    const untrusted = needsFor({ depth: "full", event: "merge_group" });
    untrusted["ci-plan"].outputs.trusted = "false";
    expect(runResult("merge_group", untrusted).exitCode).toBe(1);
  });

  test("a skipped plan passes only while the PR is still a draft", () => {
    const needs = needsFor({ depth: "fast", event: "pull_request" });
    needs["ci-plan"].result = "skipped";
    expect(runResult("pull_request", needs, "true").exitCode).toBe(0);
    expect(runResult("pull_request", needs, "false").exitCode).toBe(1);
    expect(runResult("merge_group", needs, "true").exitCode).toBe(1);
  });

  test("changed or unreadable toolchains cannot retain fast depth", () => {
    // The detector promotes both outcomes through the same validated changed output.
    const needs = needsFor({ depth: "fast", event: "pull_request", toolchainChanged: "true" });
    expect(runResult("pull_request", needs).exitCode).toBe(1);
    const full = needsFor({ depth: "full", event: "pull_request", toolchainChanged: "true" });
    expect(runResult("pull_request", full).exitCode).toBe(0);
  });

  test("toolchain promotion requires every planned area, including the API gate", () => {
    const required = planOutputs.filter((output) => output.endsWith("_required"));
    expect(required).toContain("api_required");
    expect(required).toContain("code_required");
    for (const area of required)
      for (const value of ["false", undefined]) {
        const needs = needsFor({ depth: "full", event: "pull_request", toolchainChanged: "true" });
        needs["ci-plan"].outputs[area] = value;
        expect(runResult("pull_request", needs).exitCode, `${area}=${value}`).toBe(1);
      }
  });

  test("missing or invalid detector outputs cannot approve the plan", () => {
    for (const toolchainChanged of ["", "invalid", "TRUE", null, undefined]) {
      const needs = needsFor({ depth: "fast", event: "pull_request" });
      needs["ci-plan"].outputs.toolchain_changed = toolchainChanged;
      expect(runResult("pull_request", needs).exitCode).toBe(1);
    }
  });
});
