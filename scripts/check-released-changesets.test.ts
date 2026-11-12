// A change must not add a changeset a release already published: each case
// builds a small repository with a main line, a release that deletes the
// changesets it publishes, and a change, then runs the check against it.
import { afterEach, describe, expect, setDefaultTimeout, test } from "bun:test";
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

import {
  expectExited,
  runChild,
  subprocessTestTimeout,
} from "../test/subprocess";

// The largest test here runs 17 children one after another.
setDefaultTimeout(subprocessTestTimeout(17));

const SCRIPT = path.join(import.meta.dir, "check-released-changesets.sh");

const repositories: string[] = [];

afterEach(() => {
  for (const directory of repositories.splice(0))
    rmSync(directory, { recursive: true, force: true });
});

const repository = () => {
  const directory = mkdtempSync(path.join(tmpdir(), "released-changesets-"));
  repositories.push(directory);
  const git = (...args: string[]) => {
    const run = expectExited(
      runChild(["git", ...args], {
        cwd: directory,
        env: {
          ...process.env,
          GIT_AUTHOR_NAME: "Fixture",
          GIT_AUTHOR_EMAIL: "fixture@example.com",
          GIT_COMMITTER_NAME: "Fixture",
          GIT_COMMITTER_EMAIL: "fixture@example.com",
        },
      }),
    );
    if (run.exitCode !== 0)
      throw new Error(`git ${args.join(" ")}: ${run.stderr}`);
    return run.stdout.trim();
  };
  const write = (file: string, text = "---\n---\n\nA note.\n") => {
    mkdirSync(path.dirname(path.join(directory, file)), { recursive: true });
    writeFileSync(path.join(directory, file), text);
  };
  const commit = (message: string) => {
    git("add", "--all");
    git("commit", "--quiet", "--allow-empty", "-m", message);
  };
  const check = (base = "main") => {
    const run = expectExited(
      runChild(["bash", SCRIPT, "main", base, "HEAD"], { cwd: directory }),
    );
    return { code: run.exitCode, out: run.stdout };
  };
  git("init", "--quiet", "--initial-branch=main");
  write("README.md", "fixture\n");
  commit("chore: start");
  return { git, write, commit, check };
};

/** A main line whose release published `.changeset/kept-blocks.md`. */
const releasedMain = () => {
  const repo = repository();
  repo.write(".changeset/kept-blocks.md");
  repo.commit("perf: keep blocks (#1)");
  repo.git("rm", "--quiet", ".changeset/kept-blocks.md");
  repo.write("CHANGELOG.md", "## 0.2.0\n\nKept blocks.\n");
  repo.commit("chore: version packages (#2)");
  return repo;
};

describe("released changesets", () => {
  test("a stacked branch that merged the released main fails", () => {
    const repo = repository();
    repo.git("switch", "--quiet", "-c", "stacked");
    repo.write(".changeset/kept-blocks.md");
    repo.commit("perf: keep blocks");
    repo.write(".changeset/own-note.md");
    repo.commit("perf: own change");
    repo.git("switch", "--quiet", "main");
    // The base merged with another commit: the same note, then its release.
    repo.write(".changeset/kept-blocks.md");
    repo.commit("perf: keep blocks (#1)");
    repo.git("rm", "--quiet", ".changeset/kept-blocks.md");
    repo.commit("chore: version packages (#2)");
    repo.git("switch", "--quiet", "stacked");
    repo.git("merge", "--quiet", "--no-edit", "-X", "ours", "main");
    const result = repo.check();
    expect(result.code).toBe(1);
    expect(result.out).toContain(".changeset/kept-blocks.md");
    expect(result.out).not.toContain("own-note");
  });

  test("a branch re-adding a released note fails, a new note passes", () => {
    const repo = releasedMain();
    repo.git("switch", "--quiet", "-c", "change");
    repo.write(".changeset/new-note.md");
    repo.commit("feat: new");
    expect(repo.check().code).toBe(0);
    repo.write(".changeset/kept-blocks.md");
    repo.commit("feat: same name again");
    const result = repo.check();
    expect(result.code).toBe(1);
    expect(result.out).toContain("::error file=.changeset/kept-blocks.md::");
  });

  test("only release commits count as publishing", () => {
    const repo = repository();
    repo.write(".changeset/withdrawn.md");
    repo.commit("feat: note (#1)");
    repo.git("rm", "--quiet", ".changeset/withdrawn.md");
    repo.commit("chore: withdraw a note (#2)");
    repo.git("switch", "--quiet", "-c", "change");
    repo.write(".changeset/withdrawn.md");
    repo.commit("feat: note again");
    expect(repo.check().code).toBe(0);
  });

  test("a stacked change compares with its own base", () => {
    const repo = releasedMain();
    repo.git("switch", "--quiet", "-c", "lower");
    repo.write(".changeset/lower.md");
    repo.commit("feat: lower");
    repo.git("switch", "--quiet", "-c", "upper");
    repo.write(".changeset/upper.md");
    repo.commit("feat: upper");
    expect(repo.check("lower").code).toBe(0);
    expect(repo.check("main").code).toBe(0);
  });
});
