import { afterEach, expect, setDefaultTimeout, test } from "bun:test";
import { spawnSync } from "node:child_process";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

import { subprocessTestTimeout } from "../../test/subprocess";
import { resolveComparisonBase, UNNAMED_BASE } from "./comparison-base";

// The branch test runs 17 fixture Git children and three resolver ones.
setDefaultTimeout(subprocessTestTimeout(20));

const fixtures: string[] = [];
afterEach(() => {
  for (const directory of fixtures.splice(0))
    rmSync(directory, { recursive: true, force: true });
});

const fixture = () => {
  const root = mkdtempSync(path.join(tmpdir(), "folio-comparison-base-"));
  fixtures.push(root);
  const git = (...args: string[]) => {
    const result = spawnSync(
      "git",
      [
        "-c",
        "user.name=Fixture",
        "-c",
        "user.email=fixture@example.test",
        "-c",
        "commit.gpgsign=false",
        ...args,
      ],
      { cwd: root, encoding: "utf8" },
    );
    if (result.status !== 0) throw new Error(result.stderr);
    return result.stdout.trim();
  };
  const commit = (content: string) => {
    writeFileSync(path.join(root, "file"), content);
    git("add", "file");
    git("commit", "--quiet", "-m", content);
    return git("rev-parse", "HEAD");
  };
  git("init", "--quiet", "--initial-branch=main");
  return { root, git, commit };
};

test("a hand-started run on a branch compares with where it left main", () => {
  const { root, git, commit } = fixture();
  commit("one");
  const forkPoint = commit("two");
  git("switch", "--quiet", "-c", "topic");
  commit("topic");
  git("switch", "--quiet", "main");
  const mainTip = commit("three");
  git("update-ref", "refs/remotes/origin/main", mainTip);
  git("switch", "--quiet", "topic");
  for (const requested of [UNNAMED_BASE, "", undefined])
    expect(resolveComparisonBase({ root, requested })).toBe(forkPoint);
});

test("a hand-started run on main compares with its own head", () => {
  const { root, git, commit } = fixture();
  commit("one");
  const head = commit("two");
  git("update-ref", "refs/remotes/origin/main", head);
  expect(resolveComparisonBase({ root, requested: UNNAMED_BASE })).toBe(head);
});

test("a base the event names is kept without reading the clone", () => {
  const root = mkdtempSync(path.join(tmpdir(), "folio-comparison-base-"));
  fixtures.push(root);
  const named = "0123456789abcdef0123456789abcdef01234567";
  expect(resolveComparisonBase({ root, requested: named })).toBe(named);
});

test("a clone without main fails instead of guessing a base", () => {
  const { root, commit } = fixture();
  commit("one");
  expect(() =>
    resolveComparisonBase({ root, requested: UNNAMED_BASE }),
  ).toThrow("Cannot find where HEAD left origin/main");
});

test("the workflow step prints its base as a step output", () => {
  const { root, git, commit } = fixture();
  const head = commit("one");
  git("update-ref", "refs/remotes/origin/main", head);
  const result = spawnSync(
    process.execPath,
    [path.join(import.meta.dir, "comparison-base.ts")],
    {
      cwd: root,
      encoding: "utf8",
      env: { ...process.env, TOOLCHAIN_BASE: UNNAMED_BASE },
    },
  );
  expect({ status: result.status, stdout: result.stdout }).toEqual({
    status: 0,
    stdout: `base=${head}\n`,
  });
});
