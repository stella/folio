/** Bun imports do not enforce Node JSON attributes; exercise Playwright's real loader. */
import { expect, test } from "bun:test";
import { join } from "node:path";

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const specFiles = (suite: unknown): string[] => {
  if (!isRecord(suite) || !Array.isArray(suite["specs"]))
    throw new TypeError("Playwright discovery suite has no specs.");
  const files = suite["specs"].map((spec) => {
    if (!isRecord(spec) || typeof spec["file"] !== "string")
      throw new TypeError("Playwright discovery spec has no file.");
    return spec["file"];
  });
  const children = suite["suites"];
  if (children === undefined) return files;
  if (!Array.isArray(children))
    throw new TypeError("Playwright discovery children are not suites.");
  for (const child of children) files.push(...specFiles(child));
  return files;
};

test.each([
  { project: "interactions", file: "playground-preview.interactions.spec.ts", count: 3 },
  { project: "parity-fuzzer", file: "host-api-flow.spec.ts", count: 1 },
])(
  "Node discovers $project modules and their standing regression",
  ({ project, file, count }) => {
    const root = join(import.meta.dir, "..");
    const result = Bun.spawnSync(
      [
        "node",
        join(root, "node_modules/.bin/playwright"),
        "test",
        `--project=${project}`,
        "--list",
        "--reporter=json",
      ],
      { cwd: root, stdout: "pipe", stderr: "pipe" },
    );
    const report: unknown = JSON.parse(result.stdout.toString());
    if (!isRecord(report) || !Array.isArray(report["errors"]) || !Array.isArray(report["suites"]))
      throw new TypeError(`Invalid Playwright discovery report: ${result.stderr.toString()}`);
    expect(result.exitCode, JSON.stringify(report["errors"])).toBe(0);
    expect(report["errors"]).toEqual([]);
    const files = report["suites"].flatMap(specFiles);
    expect(files.length).toBeGreaterThan(0);
    expect(files.filter((candidate) => candidate.endsWith(file))).toHaveLength(count);
  },
  30_000,
);
