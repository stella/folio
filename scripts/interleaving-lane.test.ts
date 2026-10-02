import { expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import config from "../playwright.config";

const requireRecord = (value: unknown): Record<string, unknown> => {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    throw new Error("Workflow mapping unavailable");
  }
  return value;
};

const file = "ai-human-interleaving-fuzz.interactions.spec.ts";

test("random interleaving is isolated from the PR browser and interaction projects", () => {
  const projects = config.projects;
  if (!projects) throw new Error("Playwright projects unavailable");
  const random = projects.find(({ name }) => name === "interleaving-fuzzer");
  if (!(random?.testMatch instanceof RegExp)) throw new Error("Interleaving project match missing");
  expect(random.testMatch.test(file)).toBe(true);
  const browser = projects.find(({ name }) => name === "browser-fuzzer");
  if (!(browser?.testMatch instanceof RegExp)) throw new Error("Browser project match missing");
  expect(browser.testMatch.test(file)).toBe(false);
  const interactions = projects.find(({ name }) => name === "interactions");
  if (!(interactions?.testIgnore instanceof RegExp))
    throw new Error("Interaction exclusion missing");
  expect(interactions.testIgnore.test(file)).toBe(true);
});

test("interleaving random lane uses a bounded scheduled workflow with issue filing", () => {
  const workflow = requireRecord(
    Bun.YAML.parse(readFileSync(".github/workflows/nightly-browser-input-fuzzer.yml", "utf8")),
  );
  expect(Object.keys(requireRecord(workflow["on"])).sort()).toEqual([
    "schedule",
    "workflow_dispatch",
  ]);
  const job = requireRecord(requireRecord(workflow["jobs"])["browser-input-fuzzer"]);
  const source = readFileSync(".github/workflows/nightly-browser-input-fuzzer.yml", "utf8");
  expect(source.match(/^    outputs:$/gmu)).toHaveLength(1);
  const outputs = requireRecord(job["outputs"]);
  expect(outputs["findings"]).toBe("${{ steps.fuzz.outcome == 'failure' }}");
  expect(outputs["interleaving_findings"]).toBe("${{ steps.interleaving.outcome == 'failure' }}");
  expect(job["permissions"]).toEqual({ contents: "read" });
  for (const [name, output] of [
    ["report", "findings"],
    ["interleaving-report", "interleaving_findings"],
  ]) {
    const reporter = requireRecord(requireRecord(workflow["jobs"])[name]);
    expect(reporter["if"]).toBe(
      `always() && github.ref == 'refs/heads/main' && needs.browser-input-fuzzer.outputs.${output} == 'true'`,
    );
    expect(reporter["concurrency"]).toEqual({
      group: "fuzz-failure-issues",
      "cancel-in-progress": false,
    });
  }
  const rawSteps = job["steps"];
  if (!Array.isArray(rawSteps)) throw new Error("Workflow steps unavailable");
  const steps = rawSteps.map(requireRecord);
  for (const [id, lane, spec, log] of [
    ["fuzz", "browser", "browser-input-fuzz.interactions.spec.ts", "browser-fuzz.log"],
    ["interleaving", "interleaving", file, "interleaving.log"],
  ]) {
    const run = steps.find((step) => step["id"] === id)?.["run"];
    expect(run).toContain(`--output=fuzz-playwright/${lane}`);
    expect(run).toContain(`tee fuzz-artifacts/${lane}/${log}`);
    const specSource = readFileSync(`tests/visual/${spec}`, "utf8");
    expect(specSource).toContain(`"fuzz-artifacts/${lane}/findings"`);
    const reporterName = id === "fuzz" ? "report" : "interleaving-report";
    const reporter = requireRecord(requireRecord(workflow["jobs"])[reporterName]);
    expect(JSON.stringify(reporter)).toContain(
      `--records fuzz-results/fuzz-artifacts/${lane}/findings`,
    );
    expect(JSON.stringify(reporter)).toContain(`fuzz-results/fuzz-artifacts/${lane}/${log}`);
    const artifactName = id === "fuzz" ? "browser-input-fuzzer" : "interleaving-findings";
    const upload = steps.find(
      (step) =>
        typeof step["uses"] === "string" &&
        step["uses"].startsWith("actions/upload-artifact@") &&
        requireRecord(step["with"])["name"] === artifactName,
    );
    const canonicalArtifacts = id === "fuzz" ? "fuzz-artifacts/canonical\n" : "";
    expect(requireRecord(upload?.["with"])["path"]).toBe(
      `fuzz-artifacts/${lane}\n${canonicalArtifacts}fuzz-playwright/${lane}\n`,
    );
  }
  const random = steps.find((step) => step["id"] === "interleaving");
  if (!random) throw new Error("Random workflow step unavailable");
  expect(random["continue-on-error"]).toBe(true);
  expect(random["timeout-minutes"]).toBeLessThanOrEqual(12);
  expect(random["run"]).toContain("--project=interleaving-fuzzer");
  const reporter = requireRecord(requireRecord(workflow["jobs"])["interleaving-report"]);
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

// Bind the entire workflow graph instead of mirroring the current lane list.
// New lanes must publish their step outcome and have a serialized main reporter.
test("every scheduled fuzz step has exactly one outcome output and finding reporter", () => {
  const workflow = requireRecord(
    Bun.YAML.parse(readFileSync(".github/workflows/nightly-browser-input-fuzzer.yml", "utf8")),
  );
  const jobs = requireRecord(workflow["jobs"]);
  const job = requireRecord(jobs["browser-input-fuzzer"]);
  const rawSteps = job["steps"];
  if (!Array.isArray(rawSteps)) throw new Error("Workflow steps unavailable");
  const fuzzStepIds = rawSteps
    .map(requireRecord)
    .filter((step) => step["continue-on-error"] === true)
    .map((step) => {
      const id = step["id"];
      if (typeof id !== "string") throw new Error("Fuzz step ID unavailable");
      return id;
    });
  expect(fuzzStepIds.length).toBeGreaterThan(0);
  const outputs = requireRecord(job["outputs"]);
  const outputStepIds = Object.values(outputs).map((expression) => {
    if (typeof expression !== "string") throw new Error("Fuzz output expression unavailable");
    const id = /^\$\{\{ steps\.([\w-]+)\.outcome == 'failure' \}\}$/u.exec(expression)?.at(1);
    if (!id) throw new Error("Fuzz output must report its step failure");
    return id;
  });
  expect(outputStepIds.sort()).toEqual(fuzzStepIds.sort());
  const reportedOutputs = Object.values(jobs)
    .map(requireRecord)
    .filter((candidate) => candidate["needs"] === "browser-input-fuzzer")
    .map((reporter) => {
      const condition = reporter["if"];
      if (typeof condition !== "string") throw new Error("Reporter condition unavailable");
      const output =
        /^always\(\) && github\.ref == 'refs\/heads\/main' && needs\.browser-input-fuzzer\.outputs\.(\w+) == 'true'$/u
          .exec(condition)
          ?.at(1);
      if (!output) throw new Error("Reporter must consume one outcome output on main");
      expect(reporter["concurrency"]).toEqual({
        group: "fuzz-failure-issues",
        "cancel-in-progress": false,
      });
      expect(reporter["timeout-minutes"]).toBeLessThanOrEqual(5);
      return output;
    });
  expect(reportedOutputs.sort()).toEqual(Object.keys(outputs).sort());
});
