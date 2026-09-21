import { describe, expect, test } from "bun:test";
import path from "node:path";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const PACKAGES = ["core", "react"] as const;
const PACKED_FILE_RE = /^packed\s+\S+\s+(?<file>.+)$/u;
const PRIVATE_SOURCE_RE =
  /(?:^|\/)(?:__fixtures__|__snapshots__|__tests__|fixtures)(?:\/|$)|\.test\.tsx?$/u;

const dryRunPackageFiles = (packageName: (typeof PACKAGES)[number]): string[] => {
  const result = Bun.spawnSync(["bun", "pm", "pack", "--dry-run", "--ignore-scripts"], {
    cwd: path.join(REPO_ROOT, "packages", packageName),
    stdout: "pipe",
    stderr: "pipe",
  });
  expect(result.exitCode).toBe(0);
  const output = `${result.stdout.toString()}\n${result.stderr.toString()}`;
  return output
    .split("\n")
    .map((line) => PACKED_FILE_RE.exec(line)?.groups?.["file"])
    .filter((file): file is string => file !== undefined);
};

describe("published package contents", () => {
  for (const packageName of PACKAGES) {
    test(`${packageName} excludes tests, fixtures and snapshots`, () => {
      const files = dryRunPackageFiles(packageName);
      expect(files).toContain("src/index.ts");
      expect(files.filter((file) => PRIVATE_SOURCE_RE.test(file))).toEqual([]);
    });
  }
});
