import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

const workflow = Bun.YAML.parse(
  readFileSync(fileURLToPath(new URL("../.github/workflows/ci.yml", import.meta.url)), "utf8"),
);
const resultStep = workflow.jobs["ci-result"].steps[0];
const scopes = JSON.parse(resultStep.env.JOB_SCOPES);

const needsFor = (depth, event, codeRequired = "true") => {
  const outputs = {
    trusted: "true",
    suite_depth: depth,
    code_required: codeRequired,
    typecheck_budget_required: "true",
    container_contract_required: "true",
    docx_kernel_required: "true",
    interactions_required: "true",
    browser_fuzzer_required: "true",
    differential_required: "true",
    consumer_scenarios_required: "true",
    packaged_consumer_required: "true",
  };
  const needs = { "ci-plan": { result: "success", outputs } };
  for (const [job, scope] of Object.entries(scopes)) {
    const selected =
      (scope.event === undefined || scope.event === event) &&
      (scope.area === undefined ||
        (outputs[scope.area] === (scope.value ?? "true") &&
          (scope.depth === undefined || scope.depth === depth)));
    needs[job] = { result: selected ? "success" : "skipped" };
  }
  return needs;
};

const runResult = (event, needs, draft = "false") =>
  Bun.spawnSync(["bash", "-c", `gh() { echo "$DRAFT_STATE"; }; ${resultStep.run}`], {
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
    },
  });

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
        expect(gate).toContain("github.event_name == 'pull_request'");
        expect(gate).toContain("github.event.pull_request.draft != true");
      }
      if (scope.area) {
        expect(gate).toContain(`needs.ci-plan.outputs.${scope.area} == '${scope.value ?? "true"}'`);
        expect(gate.includes("needs.ci-plan.outputs.suite_depth == 'full'")).toBe(
          scope.depth === "full",
        );
      }
    }
  });

  const scenarios = [
    ["fast", "pull_request", "true"],
    ["fast", "pull_request", "false"],
    ["fast", "push", "true"],
    ["fast", "workflow_dispatch", "true"],
    ["fast", "schedule", "true"],
    ["full", "merge_group", "true"],
  ];
  test("every declared job is selected in at least one scenario", () => {
    const exercised = new Set();
    for (const [depth, event, codeRequired] of scenarios) {
      for (const [job, { result }] of Object.entries(needsFor(depth, event, codeRequired))) {
        if (job !== "ci-plan" && result === "success") exercised.add(job);
      }
    }
    expect([...exercised].toSorted()).toEqual(Object.keys(scopes).toSorted());
  });

  for (const [depth, event, codeRequired] of scenarios) {
    test(`every selected ${depth} ${event} code=${codeRequired} job must succeed`, () => {
      const baseline = needsFor(depth, event, codeRequired);
      expect(runResult(event, baseline).exitCode).toBe(0);
      for (const [job, { result }] of Object.entries(baseline)) {
        if (job === "ci-plan" || result !== "success") continue;
        for (const failure of ["failure", "cancelled", "skipped"]) {
          const needs = structuredClone(baseline);
          needs[job].result = failure;
          const run = runResult(event, needs);
          expect(run.exitCode).toBe(1);
          expect(run.stdout.toString()).toContain(`${job} (selected, ${failure})`);
        }
      }
    });
  }

  test("an unexpected failure, cancelled plan, or wrong event depth fails", () => {
    const needs = needsFor("fast", "pull_request");
    needs["docx-kernel"].result = "failure";
    expect(runResult("pull_request", needs).exitCode).toBe(1);
    needs["ci-plan"].result = "cancelled";
    expect(runResult("pull_request", needs).exitCode).toBe(1);

    for (const event of ["pull_request", "push", "workflow_dispatch", "schedule"]) {
      for (const depth of ["full", "invalid"]) {
        const run = runResult(event, needsFor(depth, event));
        expect(run.exitCode).toBe(1);
        expect(run.stdout.toString()).toContain(`wrong suite depth for ${event}`);
      }
    }
    for (const depth of ["fast", "invalid"]) {
      const run = runResult("merge_group", needsFor(depth, "merge_group"));
      expect(run.exitCode).toBe(1);
      expect(run.stdout.toString()).toContain("wrong suite depth for merge_group");
    }

    const untrusted = needsFor("full", "merge_group");
    untrusted["ci-plan"].outputs.trusted = "false";
    expect(runResult("merge_group", untrusted).exitCode).toBe(1);
  });

  test("a skipped plan passes only while the PR is still a draft", () => {
    const needs = needsFor("fast", "pull_request");
    needs["ci-plan"].result = "skipped";
    expect(runResult("pull_request", needs, "true").exitCode).toBe(0);
    expect(runResult("pull_request", needs, "false").exitCode).toBe(1);
    expect(runResult("merge_group", needs, "true").exitCode).toBe(1);
  });
});
