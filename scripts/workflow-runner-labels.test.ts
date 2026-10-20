import { expect, setDefaultTimeout, test } from "bun:test";
import path from "node:path";
import { scanTestTimeout } from "../test/scan-timeout";

setDefaultTimeout(scanTestTimeout(10_000));

const REPO_ROOT = path.resolve(import.meta.dir, "..");

// Linux jobs name a pinned runner image. `ubuntu-latest` moves to a new image
// without a change in this repository and is served from its own runner pool,
// so the same job's tools and queue time can change between two runs.
const FLOATING_LABEL = /\bubuntu-latest\b/u;

const floatingLabelLines = (file: string, text: string): string[] =>
  text
    .split("\n")
    .flatMap((line, index) => (FLOATING_LABEL.test(line) ? [`${file}:${index + 1}`] : []));

const workflowFiles = (): readonly string[] => {
  const result = Bun.spawnSync(["git", "ls-files", ".github"], {
    cwd: REPO_ROOT,
    stdout: "pipe",
    stderr: "pipe",
  });
  expect(result.exitCode).toBe(0);
  return result.stdout
    .toString()
    .split("\n")
    .filter((file) => /\.ya?ml$/u.test(file));
};

test("no workflow or action names the floating Linux label", async () => {
  const files = workflowFiles();
  expect(files.length).toBeGreaterThan(0);
  const offenders: string[] = [];
  for (const file of files) {
    offenders.push(...floatingLabelLines(file, await Bun.file(path.join(REPO_ROOT, file)).text()));
  }
  expect(offenders).toEqual([]);
});

test("a job on the floating label is reported with its line", () => {
  const workflow =
    "jobs:\n  build:\n    runs-on: ubuntu-latest\n  pinned:\n    runs-on: ubuntu-24.04\n";
  expect(floatingLabelLines(".github/workflows/example.yml", workflow)).toEqual([
    ".github/workflows/example.yml:3",
  ]);
});
