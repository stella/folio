import { expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import config from "../playwright.config";

const requireRecord = (value: unknown): Record<string, unknown> => {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    throw new Error("Workflow mapping unavailable");
  }
  return value;
};

const specs = [
  "measure-backend-parity.spec.ts",
  "engine-layout-parity.spec.ts",
  "measure-parity.spec.ts",
  "interactions.spec.ts",
  "rendering.spec.ts",
];

test("second-engine specs run only in the report-only projects", () => {
  const projects = config.projects;
  if (!projects) throw new Error("Playwright projects unavailable");
  const matching = (file: string) =>
    projects
      .filter(({ testMatch }) => testMatch instanceof RegExp && testMatch.test(file))
      .map(({ name }) => name);
  expect(matching("engine-layout-parity.spec.ts").toSorted()).toEqual([
    "engine-layout-record",
    "webkit-layout-parity",
  ]);
  expect(matching("measure-backend-parity.spec.ts").toSorted()).toEqual([
    "measure-parity",
    "webkit-measure-backend",
  ]);
  for (const file of specs.slice(2)) {
    expect(matching(file)).not.toContain("webkit-layout-parity");
  }
  const webkit = projects.find(({ name }) => name === "webkit-layout-parity");
  expect(webkit?.use?.browserName).toBe("webkit");
  const record = projects.find(({ name }) => name === "engine-layout-record");
  expect(record?.use?.browserName).toBeUndefined();
});

test("the WebKit workflow is nightly, on macOS, read-only and never files issues", () => {
  const source = readFileSync(".github/workflows/nightly-webkit-parity.yml", "utf8");
  const workflow = requireRecord(Bun.YAML.parse(source));
  expect(Object.keys(requireRecord(workflow["on"])).toSorted()).toEqual([
    "schedule",
    "workflow_dispatch",
  ]);
  const jobs = requireRecord(workflow["jobs"]);
  expect(Object.keys(jobs)).toEqual(["webkit-layout-parity"]);
  const job = requireRecord(jobs["webkit-layout-parity"]);
  expect(job["runs-on"]).toBe("macos-latest");
  expect(job["permissions"]).toEqual({ contents: "read" });
  for (const project of [
    "engine-layout-record",
    "webkit-layout-parity",
    "webkit-measure-backend",
  ]) {
    expect(source).toContain(`--project=${project}`);
  }
  expect(source).toContain("playwright install chromium webkit");
  expect(source).not.toContain("issues:");
  expect(JSON.stringify(readFileSync(".github/workflows/ci.yml", "utf8"))).not.toContain(
    "webkit-layout-parity",
  );
});
