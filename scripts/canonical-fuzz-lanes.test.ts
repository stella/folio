/** Random canonical flows run explicitly at night and cannot enter required CI. */
import { expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";
import ts from "typescript";

import playwright from "../playwright.config";

const ROOT = path.resolve(import.meta.dir, "..");
const PROPERTY_FILE = "test/canonical-session.fuzz.ts";
const BROWSER_FILE = "canonical-browser-input-fuzz.interactions.spec.ts";
const replayCallDefects = (text: string) => {
  const source = ts.createSourceFile("replay.ts", text, ts.ScriptTarget.Latest, true);
  const required = new Set<string>();
  const calls: ts.CallExpression[] = [];
  const visit = (node: ts.Node): void => {
    if (ts.isTypeAliasDeclaration(node) && node.name.text === "RunFuzzOptions") {
      if (!ts.isTypeLiteralNode(node.type)) throw new TypeError("Missing replay options fields.");
      for (const member of node.type.members) {
        if (ts.isPropertySignature(member) && !member.questionToken)
          required.add(member.name.getText(source));
      }
    }
    if (ts.isCallExpression(node) && node.expression.getText(source) === "runMode")
      calls.push(node);
    ts.forEachChild(node, visit);
  };
  visit(source);
  if (required.size === 0 || calls.length === 0) throw new TypeError("Missing replay contract.");
  return calls.flatMap((call) => {
    const options = call.arguments.at(0);
    if (call.arguments.length !== 1 || !options || !ts.isObjectLiteralExpression(options))
      return [call.getText(source)];
    const keys = new Set(
      options.properties.flatMap((property) =>
        ts.isPropertyAssignment(property) || ts.isShorthandPropertyAssignment(property)
          ? [property.name.getText(source)]
          : [],
      ),
    );
    return [...required].every((key) => keys.has(key)) ? [] : [call.getText(source)];
  });
};

test("every browser replay caller supplies the declared options contract", () => {
  expect(
    replayCallDefects(
      readFileSync(path.join(ROOT, "tests/visual/browser-input-fuzz.interactions.spec.ts"), "utf8"),
    ),
  ).toEqual([]);
  const contract = "type RunFuzzOptions = { page: Page; authority: Authority };";
  expect(replayCallDefects(`${contract} runMode(page, source);`)).toHaveLength(1);
  expect(replayCallDefects(`${contract} runMode({ page });`)).toHaveLength(1);
  expect(replayCallDefects(`${contract} runMode({ page, authority });`)).toEqual([]);
});
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

const requiredCiSteps = (readWorkflow: (name: string) => Record<string, unknown>) => {
  const pending = ["ci.yml"];
  const visited = new Set<string>();
  const steps: Record<string, unknown>[] = [];
  while (pending.length > 0) {
    const name = pending.pop();
    if (name === undefined || visited.has(name)) continue;
    visited.add(name);
    const jobs = readWorkflow(name)["jobs"];
    if (!record(jobs)) throw new Error(`Missing jobs in ${name}`);
    for (const job of Object.values(jobs)) {
      if (!record(job)) throw new Error(`Invalid job in ${name}`);
      if (typeof job["uses"] !== "string") {
        steps.push(...stepsOf(job));
        continue;
      }
      const prefix = "./.github/workflows/";
      const file = job["uses"].slice(prefix.length);
      if (!job["uses"].startsWith(prefix) || file === "" || path.posix.basename(file) !== file)
        throw new Error(`Cannot inspect reusable workflow ${job["uses"]}`);
      pending.push(file);
    }
  }
  return steps;
};

const canonicalHarnessCalls = (steps: Record<string, unknown>[]) =>
  steps.flatMap((step) => {
    const run = String(step["run"] ?? "");
    return /canonical-session|--project=browser-fuzzer/u.test(run) ? [run] : [];
  });

test("required CI never invokes either random canonical harness", () => {
  expect(canonicalHarnessCalls(requiredCiSteps(workflow))).toEqual([]);
  // Bun and the property-area collector discover .test/.spec, not .fuzz files.
  expect(PROPERTY_FILE).not.toMatch(/(?:\.test|_test_|\.spec|_spec_)/);
  expect(readFileSync(path.join(ROOT, PROPERTY_FILE), "utf8")).toContain("assertProperty(");
});

test("the required CI guard inspects nested reusable workflows", () => {
  for (const run of [
    "bun test test/canonical-session.fuzz.ts",
    "bunx playwright test --project=browser-fuzzer",
  ]) {
    const fixtures: Record<string, Record<string, unknown>> = {
      "ci.yml": { jobs: { first: { uses: "./.github/workflows/first.yml" } } },
      "first.yml": { jobs: { nested: { uses: "./.github/workflows/second.yml" } } },
      "second.yml": { jobs: { runner: { steps: [{ run }] } } },
    };
    const readFixture = (name: string) => {
      const fixture = fixtures[name];
      if (fixture === undefined) throw new Error(`Missing fixture ${name}`);
      return fixture;
    };
    expect(canonicalHarnessCalls(requiredCiSteps(readFixture))).toEqual([run]);
    fixtures["second.yml"] = { jobs: { runner: { steps: [{ run: "bun run test:smoke" }] } } };
    expect(canonicalHarnessCalls(requiredCiSteps(readFixture))).toEqual([]);
  }
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
