import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import {
  existsSync,
  mkdtempSync,
  mkdirSync,
  readFileSync,
  rmSync,
  statSync,
  utimesSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, test } from "bun:test";

import { ensurePlaygroundBuild } from "./playground-build";

const temporaryRoots: string[] = [];

const createRepository = (): { repoRoot: string; packageDir: string } => {
  const repoRoot = mkdtempSync(join(tmpdir(), "folio-playground-build-"));
  temporaryRoots.push(repoRoot);
  execFileSync("git", ["init", "--quiet"], { cwd: repoRoot });
  const packageDir = join(repoRoot, "packages", "playground");
  mkdirSync(packageDir, { recursive: true });
  writeFileSync(
    join(packageDir, "package.json"),
    JSON.stringify({ name: "@stll/playground", scripts: { build: "vite build" } }),
  );
  writeFileSync(join(repoRoot, ".gitignore"), "dist/\npackages/playground/.env.production.local\n");
  writeFileSync(join(packageDir, "index.html"), "<main>fixture</main>");
  execFileSync("git", ["add", "-A"], { cwd: repoRoot });
  return { repoRoot, packageDir };
};

const writeBuildOutput = (packageDir: string, content = "bundle"): void => {
  const assetsDir = join(packageDir, "dist", "assets");
  mkdirSync(assetsDir, { recursive: true });
  writeFileSync(join(packageDir, "dist", "index.html"), "<main>built</main>");
  writeFileSync(join(assetsDir, "index.js"), content);
};

const buildWithOutput =
  (packageDir: string, content = "bundle") =>
  () =>
    writeBuildOutput(packageDir, content);

afterEach(() => {
  for (const root of temporaryRoots.splice(0)) rmSync(root, { recursive: true, force: true });
});

describe("incremental playground build cache", () => {
  test("invalidates on added, edited, and deleted tracked inputs", async () => {
    const { repoRoot, packageDir } = createRepository();
    let builds = 0;
    const build = () => {
      builds += 1;
      writeBuildOutput(packageDir);
    };
    const ensure = () => ensurePlaygroundBuild({ repoRoot, packageDir, build, log: () => {} });

    assert.equal((await ensure()).status, "built");
    assert.equal((await ensure()).status, "skipped");
    assert.equal(builds, 1);

    writeFileSync(join(packageDir, "added.ts"), "export const added = true;");
    assert.equal((await ensure()).status, "built");

    const addedInput = join(packageDir, "added.ts");
    writeFileSync(addedInput, "export const added = 'edited';");
    const editedTime = new Date(statSync(addedInput).mtimeMs + 2_000);
    utimesSync(addedInput, editedTime, editedTime);
    assert.equal((await ensure()).status, "built");

    rmSync(addedInput);
    assert.equal((await ensure()).status, "built");
    rmSync(join(packageDir, "index.html"));
    assert.equal((await ensure()).status, "built");
    assert.equal(builds, 5);
  });

  test("includes ignored production Vite env files and invalidates missing or changed output files", async () => {
    const { repoRoot, packageDir } = createRepository();
    const envPath = join(packageDir, ".env.production.local");
    writeFileSync(envPath, "VITE_FEATURE=one");
    execFileSync("git", ["check-ignore", "--quiet", envPath], { cwd: repoRoot, stdio: "ignore" });
    const ensureBuild = (build: () => void) =>
      ensurePlaygroundBuild({ repoRoot, packageDir, build, log: () => {} });

    assert.equal((await ensureBuild(buildWithOutput(packageDir))).status, "built");
    assert.equal((await ensureBuild(buildWithOutput(packageDir))).status, "skipped");

    writeFileSync(envPath, "VITE_FEATURE=changed");
    const editedTime = new Date(statSync(envPath).mtimeMs + 2_000);
    utimesSync(envPath, editedTime, editedTime);
    assert.equal((await ensureBuild(buildWithOutput(packageDir))).status, "built");

    const output = join(packageDir, "dist", "assets", "index.js");
    rmSync(output);
    assert.equal((await ensureBuild(buildWithOutput(packageDir))).status, "built");

    writeFileSync(output, "changed-output-size");
    const outputTime = new Date(statSync(output).mtimeMs + 2_000);
    utimesSync(output, outputTime, outputTime);
    assert.equal((await ensureBuild(buildWithOutput(packageDir))).status, "built");
  });

  test("fingerprints NODE_ENV and VITE build variables without storing their values", async () => {
    const { repoRoot, packageDir } = createRepository();
    const ensureWithEnvironment = (buildEnvironment: Record<string, string | undefined>) =>
      ensurePlaygroundBuild({
        repoRoot,
        packageDir,
        build: buildWithOutput(packageDir),
        buildEnvironment,
        log: () => {},
      });

    assert.equal(
      (await ensureWithEnvironment({ NODE_ENV: "production", VITE_API_URL: "https://one.test" }))
        .status,
      "built",
    );
    assert.equal(
      (await ensureWithEnvironment({ NODE_ENV: "production", VITE_API_URL: "https://one.test" }))
        .status,
      "skipped",
    );
    assert.equal(
      (await ensureWithEnvironment({ NODE_ENV: "production", VITE_API_URL: "https://two.test" }))
        .status,
      "built",
    );
    const stamp = readFileSync(join(packageDir, "dist", ".playground-build-cache.json"), "utf8");
    assert.equal(stamp.includes("https://two.test"), false);
    assert.equal(
      (await ensureWithEnvironment({ NODE_ENV: "development", VITE_API_URL: "https://two.test" }))
        .status,
      "built",
    );
  });

  test("rebuilds for a corrupt cache and a missing entry page", async () => {
    const { repoRoot, packageDir } = createRepository();
    const ensure = () =>
      ensurePlaygroundBuild({
        repoRoot,
        packageDir,
        build: buildWithOutput(packageDir),
        log: () => {},
      });

    assert.equal((await ensure()).status, "built");
    writeFileSync(join(packageDir, "dist", ".playground-build-cache.json"), "{");
    assert.equal((await ensure()).status, "built");

    rmSync(join(packageDir, "dist", "index.html"));
    assert.equal((await ensure()).status, "built");
    assert.equal(existsCacheStamp(packageDir), true);
  });

  test("a failed build leaves no cache stamp", async () => {
    const { repoRoot, packageDir } = createRepository();
    const ensure = (build: () => void) =>
      ensurePlaygroundBuild({ repoRoot, packageDir, build, log: () => {} });

    await assert.rejects(
      ensure(() => {
        writeBuildOutput(packageDir);
        throw new TypeError("fixture build failed");
      }),
      /fixture build failed/u,
    );
    assert.equal(existsCacheStamp(packageDir), false);

    assert.equal((await ensure(buildWithOutput(packageDir))).status, "built");
    assert.equal(existsCacheStamp(packageDir), true);
  });

  test("does not cache output when inputs change during the build", async () => {
    const { repoRoot, packageDir } = createRepository();
    const input = join(packageDir, "index.html");
    await assert.rejects(
      ensurePlaygroundBuild({
        repoRoot,
        packageDir,
        build: () => {
          writeBuildOutput(packageDir);
          writeFileSync(input, "<main>changed during build</main>");
          const changedTime = new Date(statSync(input).mtimeMs + 2_000);
          utimesSync(input, changedTime, changedTime);
        },
        log: () => {},
      }),
      /inputs or build environment changed/u,
    );
    assert.equal(existsCacheStamp(packageDir), false);
    assert.equal(
      (
        await ensurePlaygroundBuild({
          repoRoot,
          packageDir,
          build: buildWithOutput(packageDir),
          log: () => {},
        })
      ).status,
      "built",
    );
  });

  test("rejects build-environment changes during the build without writing a stamp", async () => {
    const { repoRoot, packageDir } = createRepository();
    const buildEnvironment: Record<string, string | undefined> = {
      NODE_ENV: "production",
      VITE_API_URL: "https://before.test",
    };
    await assert.rejects(
      ensurePlaygroundBuild({
        repoRoot,
        packageDir,
        buildEnvironment,
        build: () => {
          writeBuildOutput(packageDir);
          buildEnvironment["VITE_API_URL"] = "https://after.test";
        },
        log: () => {},
      }),
      /inputs or build environment changed/u,
    );
    assert.equal(existsCacheStamp(packageDir), false);
    assert.equal(
      (
        await ensurePlaygroundBuild({
          repoRoot,
          packageDir,
          build: buildWithOutput(packageDir),
          buildEnvironment,
          log: () => {},
        })
      ).status,
      "built",
    );
  });
});

const existsCacheStamp = (packageDir: string): boolean => {
  return existsSync(join(packageDir, "dist", ".playground-build-cache.json"));
};
