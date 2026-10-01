import { expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { validateDocxPackage } from "../packages/docx-core/src/validate/docx";
import config from "../playwright.config";

const requireRecord = (value: unknown): Record<string, unknown> => {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    throw new Error("Workflow mapping unavailable");
  }
  return value;
};

const file = "cross-host-flow.spec.ts";

test("cross-host validator resolves from its owning source without a root package dependency", async () => {
  const specPath = resolve("tests/parity", file);
  const source = readFileSync(specPath, "utf8");
  const validatorImport = source
    .match(/import\s*\{\s*validateDocxPackage\s*\}\s*from\s*["']([^"']+)["']/)
    ?.at(1);
  if (!validatorImport) throw new Error("Cross-host validator import unavailable");
  expect(validatorImport.startsWith(".")).toBe(true);
  expect(Bun.resolveSync(validatorImport, dirname(specPath))).toBe(
    resolve("packages/docx-core/src/validate/docx.ts"),
  );
  expect(await validateDocxPackage(new Uint8Array())).toMatchObject({
    valid: false,
    code: "invalid_archive",
  });
});

test("random cross-host flow is isolated from deterministic parity and Vue projects", () => {
  const projects = config.projects;
  if (!projects) throw new Error("Playwright projects unavailable");
  const random = projects.find(({ name }) => name === "parity-fuzzer");
  if (!(random?.testMatch instanceof RegExp))
    throw new Error("Random parity project match missing");
  expect(random.testMatch.test(file)).toBe(true);
  for (const name of ["parity", "vue"]) {
    const project = projects.find((candidate) => candidate.name === name);
    if (!(project?.testIgnore instanceof RegExp)) throw new Error(`${name} exclusion missing`);
    expect(project.testIgnore.test(file)).toBe(true);
  }
});

test("cross-host random lane uses a bounded scheduled workflow with issue filing", () => {
  const workflow = requireRecord(
    Bun.YAML.parse(readFileSync(".github/workflows/nightly-parity-e2e.yml", "utf8")),
  );
  expect(Object.keys(requireRecord(workflow["on"])).sort()).toEqual([
    "schedule",
    "workflow_dispatch",
  ]);
  const job = requireRecord(requireRecord(workflow["jobs"])["parity-e2e"]);
  const rawSteps = job["steps"];
  if (!Array.isArray(rawSteps)) throw new Error("Workflow steps unavailable");
  const steps = rawSteps.map(requireRecord);
  const random = steps.find((step) => step["id"] === "cross-host");
  if (!random) throw new Error("Random workflow step unavailable");
  expect(random["continue-on-error"]).toBe(true);
  expect(random["timeout-minutes"]).toBeLessThanOrEqual(12);
  expect(random["run"]).toContain("--project=parity-fuzzer");
  const reporter = requireRecord(requireRecord(workflow["jobs"])["cross-host-report"]);
  expect(reporter["timeout-minutes"]).toBeLessThanOrEqual(5);
  const reportSteps = reporter["steps"];
  if (!Array.isArray(reportSteps)) throw new Error("Report steps unavailable");
  expect(
    reportSteps
      .map(requireRecord)
      .some(
        (step) =>
          typeof step["run"] === "string" && step["run"].includes("scripts/fuzz-failure-issues.ts"),
      ),
  ).toBe(true);
});
