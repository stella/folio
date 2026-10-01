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
  expect(random["run"]).toContain("--output=fuzz-playwright/cross-host");
  expect(random["run"]).toContain("tee fuzz-artifacts/cross-host/cross-host.log");
  const deterministic = steps.find((step) => step["name"] === "Parity e2e specs");
  expect(deterministic?.["run"]).toContain("--output=fuzz-playwright/parity");
  const spec = readFileSync(resolve("tests/parity", file), "utf8");
  const recordsPath = spec.match(/writeFailureRecord\(\s*"([^"]+)"/)?.at(1);
  expect(recordsPath).toBe("fuzz-artifacts/cross-host/findings");
  const upload = steps.find((step) => step["name"] === "Upload differential repro artifacts");
  if (!upload) throw new Error("Repro artifact upload unavailable");
  expect(requireRecord(upload["with"])["path"]).toBe("fuzz-artifacts/\nfuzz-playwright/\n");
  const reporter = requireRecord(requireRecord(workflow["jobs"])["cross-host-report"]);
  expect(reporter["timeout-minutes"]).toBeLessThanOrEqual(5);
  expect(reporter["if"]).toContain("github.ref == 'refs/heads/main'");
  const reportSteps = reporter["steps"];
  if (!Array.isArray(reportSteps)) throw new Error("Report steps unavailable");
  const report = reportSteps
    .map(requireRecord)
    .find(
      (step) =>
        typeof step["run"] === "string" && step["run"].includes("scripts/fuzz-failure-issues.ts"),
    );
  expect(report?.["run"]).toContain("--log fuzz-results/fuzz-artifacts/cross-host/cross-host.log");
  expect(report?.["run"]).toContain(`--records fuzz-results/${recordsPath}`);
});
