import { afterEach, expect, test } from "bun:test";
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

import { affectedPackagePlan, changedWorkingTreePaths } from "./ci-api-plan";

const fixtures = new Set<string>();
afterEach(() => {
  for (const directory of fixtures)
    rmSync(directory, { recursive: true, force: true });
  fixtures.clear();
});

const plan = (changedPaths: string[]) =>
  affectedPackagePlan({
    changedPaths,
    publishedPackagePaths: ["packages/public"],
  });

test("published sources, consumer scenarios, and docs select the expected checks", () => {
  expect(plan(["packages/public/src/index.ts"])).toEqual({
    apiRequired: true,
    packagedConsumerRequired: true,
  });
  expect(plan(["test/consumer-scenarios/scenarios/compare.test.ts"])).toEqual({
    apiRequired: false,
    packagedConsumerRequired: true,
  });
  expect(plan(["docs/guide.md"])).toEqual({
    apiRequired: false,
    packagedConsumerRequired: false,
  });
});

test("working-tree paths include committed, staged, unstaged, and untracked files", () => {
  const repoRoot = mkdtempSync(path.join(tmpdir(), "folio-affected-packages-"));
  fixtures.add(repoRoot);
  const git = (...args: string[]) => {
    const result = Bun.spawnSync(["git", ...args], {
      cwd: repoRoot,
      stdout: "pipe",
      stderr: "pipe",
    });
    if (result.exitCode !== 0) throw new TypeError(result.stderr.toString());
    return result.stdout.toString().trim();
  };
  const write = (file: string, contents: string) => {
    const destination = path.join(repoRoot, file);
    mkdirSync(path.dirname(destination), { recursive: true });
    writeFileSync(destination, contents);
  };
  git("init", "--quiet");
  git("config", "user.email", "fixture@example.invalid");
  git("config", "user.name", "Fixture");
  write("tracked.ts", "initial\n");
  git("add", ".");
  git("commit", "--quiet", "--no-gpg-sign", "-m", "initial");
  const base = git("rev-parse", "HEAD");
  write("committed.ts", "committed\n");
  git("add", ".");
  git("commit", "--quiet", "--no-gpg-sign", "-m", "next");
  write("staged.ts", "staged\n");
  git("add", "staged.ts");
  write("tracked.ts", "unstaged\n");
  write("untracked.ts", "untracked\n");
  expect(changedWorkingTreePaths({ repoRoot, base })).toEqual([
    "committed.ts",
    "staged.ts",
    "tracked.ts",
    "untracked.ts",
  ]);
});
