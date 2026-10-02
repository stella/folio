/** Random canonical flows run explicitly at night and cannot enter required CI. */
import { expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";

import playwright from "../playwright.config";

const ROOT = path.resolve(import.meta.dir, "..");
const PROPERTY_FILE = "test/canonical-session.fuzz.ts";
const BROWSER_FILE = "canonical-browser-input-fuzz.interactions.spec.ts";
const record = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);
const workflow = (name: string): Record<string, unknown> => {
  const value: unknown = Bun.YAML.parse(
    readFileSync(path.join(ROOT, ".github/workflows", name), "utf8"),
  );
  if (!record(value)) throw new Error(`Invalid workflow ${name}`);
  return value;
};

const stepsOf = (job: unknown): Record<string, unknown>[] => {
  if (!record(job) || !Array.isArray(job["steps"])) throw new Error("Missing workflow steps");
  return job["steps"].filter(record);
};

test("required CI never invokes either random canonical harness", () => {
  const jobs = workflow("ci.yml")["jobs"];
  if (!record(jobs)) throw new Error("Missing CI jobs");
  for (const job of Object.values(jobs)) {
    for (const step of stepsOf(job)) {
      expect(String(step["run"] ?? "")).not.toMatch(/canonical-session|--project=browser-fuzzer/);
    }
  }
  // Bun and the property-area collector discover .test/.spec, not .fuzz files.
  expect(PROPERTY_FILE).not.toMatch(/(?:\.test|_test_|\.spec|_spec_)/);
  expect(readFileSync(path.join(ROOT, PROPERTY_FILE), "utf8")).toContain("assertProperty(");
});

test("nightly explicitly executes both lanes and keeps logs for issue filing", () => {
  const nightly = workflow("nightly-browser-input-fuzzer.yml");
  const triggers = nightly["on"];
  const jobs = nightly["jobs"];
  if (!record(triggers) || !record(jobs)) throw new Error("Missing nightly workflow structure");
  expect(Object.keys(triggers).toSorted()).toEqual(["schedule", "workflow_dispatch"]);
  const fuzz = jobs["browser-input-fuzzer"];
  if (!record(fuzz)) throw new Error("Missing nightly fuzz job");
  expect(fuzz["permissions"]).toEqual({ contents: "read" });
  const steps = stepsOf(fuzz);
  const property = steps.find((step) => step["id"] === "canonical");
  const browser = steps.find((step) => step["id"] === "fuzz");
  expect(property?.["run"]).toContain(`bun test ./${PROPERTY_FILE}`);
  expect(property?.["run"]).toContain("tee fuzz-artifacts/canonical/canonical-fuzz.log");
  expect(browser?.["run"]).toContain("--project=browser-fuzzer");
  expect(browser?.["run"]).toContain("tee fuzz-artifacts/browser/browser-fuzz.log");
  expect(
    steps.find((step) => step["name"] === "Preserve nightly failure status")?.["if"],
  ).toContain("steps.canonical.outcome == 'failure'");
  const report = jobs["report-canonical"];
  if (!record(report)) throw new Error("Missing issue reporter");
  expect(report["permissions"]).toEqual({ contents: "read", issues: "write" });
  expect(report["if"]).toContain("github.ref == 'refs/heads/main'");
  expect(report["concurrency"]).toEqual({
    group: "fuzz-failure-issues",
    "cancel-in-progress": false,
  });
  const filing = stepsOf(report).find((step) =>
    String(step["run"] ?? "").includes("scripts/fuzz-failure-issues.ts"),
  );
  expect(filing?.["run"]).toContain("fuzz-results/fuzz-artifacts/canonical/canonical-fuzz.log");
  const browserReport = jobs["report"];
  if (!record(browserReport)) throw new Error("Missing browser issue reporter");
  const browserFiling = stepsOf(browserReport).find((step) =>
    String(step["run"] ?? "").includes("scripts/fuzz-failure-issues.ts"),
  );
  expect(browserFiling?.["run"]).toContain("fuzz-results/fuzz-artifacts/browser/browser-fuzz.log");
  expect(report["if"]).toContain("outputs.canonical_findings == 'true'");
  expect(filing?.["run"]).toContain("--records fuzz-results/fuzz-artifacts/canonical/findings");
});

test("all nightly evidence survives Playwright output cleanup and is uploaded", () => {
  const jobs = workflow("nightly-browser-input-fuzzer.yml")["jobs"];
  if (!record(jobs)) throw new Error("Missing nightly jobs");
  const steps = stepsOf(jobs["browser-input-fuzzer"]);
  const output = path.resolve(ROOT, playwright.outputDir ?? "test-results");
  const lanes = steps.filter((step) => step["id"] === "canonical" || step["id"] === "fuzz");
  expect(lanes).toHaveLength(2);
  for (const lane of lanes) {
    const env = lane["env"];
    if (!record(env)) throw new Error("Missing lane environment");
    for (const [key, value] of Object.entries(env)) {
      if (!key.endsWith("FAILURES_DIR")) continue;
      expect(path.resolve(ROOT, String(value)).startsWith(`${output}${path.sep}`)).toBe(false);
      expect(String(value)).toBe("fuzz-artifacts/canonical/findings");
    }
    expect(lane["continue-on-error"]).toBe(true);
    expect(lane["run"]).not.toContain("tee test-results/");
  }
  const fuzz = jobs["browser-input-fuzzer"];
  if (!record(fuzz) || !record(fuzz["outputs"])) throw new Error("Missing fuzz outputs");
  expect(fuzz["outputs"]["canonical_findings"]).toBe("${{ steps.canonical.outcome == 'failure' }}");
  const upload = steps.find((step) => step["name"] === "Keep health log and minimized traces");
  const options = upload?.["with"];
  if (!record(options)) throw new Error("Missing upload options");
  expect(String(options["path"]).split("\n")).toContain("fuzz-artifacts/canonical");
  for (const file of [BROWSER_FILE, "browser-input-fuzz.interactions.spec.ts"]) {
    expect(readFileSync(path.join(ROOT, "tests/visual", file), "utf8")).toContain(
      'process.env["FOLIO_FUZZ_FAILURES_DIR"]',
    );
  }
});

test("the interactions project excludes the canonical browser fuzzer", () => {
  const exclusion = playwright.projects?.find(
    (project) => project.name === "interactions",
  )?.testIgnore;
  const match = playwright.projects?.find(
    (project) => project.name === "browser-fuzzer",
  )?.testMatch;
  if (!(exclusion instanceof RegExp) || !(match instanceof RegExp))
    throw new Error("Missing browser project matchers");
  expect(exclusion.test(BROWSER_FILE)).toBe(true);
  expect(match.test(BROWSER_FILE)).toBe(true);
});
