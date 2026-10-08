/** Workspace and Git fixtures enumerate the API gate's source and control boundaries. */
import { afterEach, describe, expect, test } from "bun:test";
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { apiRequiredForPaths, changedApiPaths, publishedPackagePaths } from "./ci-api-plan";

const fixtures = new Set<string>();
afterEach(() => {
  for (const fixture of fixtures) rmSync(fixture, { recursive: true, force: true });
  fixtures.clear();
});

const write = (repoRoot: string, file: string, contents: string) => {
  const destination = path.join(repoRoot, file);
  mkdirSync(path.dirname(destination), { recursive: true });
  writeFileSync(destination, contents);
};
type FixtureOptions = { workspaces: unknown; manifests: Readonly<Record<string, unknown>> };
const fixture = ({ workspaces, manifests }: FixtureOptions) => {
  const repoRoot = mkdtempSync(path.join(tmpdir(), "folio-ci-api-"));
  fixtures.add(repoRoot);
  write(
    repoRoot,
    "package.json",
    JSON.stringify({ name: "fixture-root", private: true, workspaces }),
  );
  for (const [directory, manifest] of Object.entries(manifests))
    write(repoRoot, `${directory}/package.json`, JSON.stringify(manifest));
  return repoRoot;
};

type GitOptions = { repoRoot: string; args: string[]; input?: string };
const git = ({ repoRoot, args, input }: GitOptions) => {
  const result = Bun.spawnSync(["git", ...args], {
    cwd: repoRoot,
    stdout: "pipe",
    stderr: "pipe",
    ...(input === undefined ? {} : { stdin: Buffer.from(input) }),
  });
  if (result.exitCode !== 0) throw new TypeError(result.stderr.toString());
  return result.stdout.toString().trim();
};
// Synthetic commit objects require neither a user identity nor commit hooks/signing.
const snapshot = (repoRoot: string, parent?: string) => {
  git({ repoRoot, args: ["add", "--", "."] });
  const tree = git({ repoRoot, args: ["write-tree"] });
  const parents = parent === undefined ? "" : `parent ${parent}\n`;
  const input = `tree ${tree}\n${parents}author Fixture <fixture@example.invalid> 1700000000 +0000\ncommitter Fixture <fixture@example.invalid> 1700000000 +0000\n\nfixture\n`;
  return git({ repoRoot, args: ["hash-object", "-t", "commit", "-w", "--stdin"], input });
};
const gitFixture = () => {
  const repoRoot = fixture({
    workspaces: ["modules/*"],
    manifests: {
      "modules/public": { name: "public-package" },
      "modules/private": { name: "private-package", private: true },
    },
  });
  git({ repoRoot, args: ["init", "--quiet"] });
  write(repoRoot, "modules/public/src/original.ts", "export const original = 1;\n");
  write(repoRoot, "modules/private/src/private.ts", "export const privateValue = 1;\n");
  write(repoRoot, "docs/guide.md", "guide\n");
  return { repoRoot, base: snapshot(repoRoot) };
};

describe("manifest-derived published API scope", () => {
  test("every currently published workspace selects its source tree", () => {
    const published = publishedPackagePaths();
    expect(published.length).toBeGreaterThan(0);
    for (const directory of published)
      expect(
        apiRequiredForPaths({
          changedPaths: [`${directory}/src/fixture.ts`],
          publishedPackagePaths: published,
        }),
      ).toBe(true);
  });
  test("root globs discover each non-private workspace, deduplicate overlaps and honor exclusions", () => {
    const repoRoot = fixture({
      workspaces: ["modules/*", "modules/open", "plugins/deep/*", "!modules/excluded"],
      manifests: {
        "modules/open": { name: "implicit-public" },
        "modules/explicit": { name: "explicit-public", private: false },
        "modules/private": { name: "private", private: true },
        "modules/excluded": { name: "excluded" },
        "plugins/deep/adapter": { name: "nested-public" },
        "outside/unlisted": { name: "not-a-workspace" },
      },
    });
    expect(publishedPackagePaths(repoRoot)).toEqual([
      "modules/explicit",
      "modules/open",
      "plugins/deep/adapter",
    ]);
    write(repoRoot, "modules/new-package/package.json", JSON.stringify({ name: "new-public" }));
    const published = publishedPackagePaths(repoRoot);
    expect(published).toContain("modules/new-package");
    for (const directory of published)
      for (const suffix of [
        "entry.ts",
        "deep/types.d.ts",
        "schema.json",
        "assets/file with spaces.bin",
      ])
        expect(
          apiRequiredForPaths({
            changedPaths: [`${directory}/src/${suffix}`],
            publishedPackagePaths: published,
          }),
        ).toBe(true);
  });

  test("every newly discovered published compiler config selects API checks", () => {
    const repoRoot = fixture({
      workspaces: ["modules/*"],
      manifests: { "modules/first": { name: "first" } },
    });
    write(repoRoot, "modules/new-package/package.json", JSON.stringify({ name: "new-package" }));
    const published = publishedPackagePaths(repoRoot);
    expect(published).toContain("modules/new-package");
    for (const directory of published)
      expect(
        apiRequiredForPaths({
          changedPaths: [`${directory}/tsconfig.json`],
          publishedPackagePaths: published,
        }),
      ).toBe(true);
  });

  test("docs and private source edits do not select API; every API control does", () => {
    const repoRoot = fixture({
      workspaces: ["modules/*"],
      manifests: {
        "modules/public": { name: "public" },
        "modules/private": { name: "private", private: true },
      },
    });
    const published = publishedPackagePaths(repoRoot);
    const outside = [
      "docs/guide.md",
      "README.md",
      "modules/public/README.md",
      "modules/public/src-other/file.ts",
      "modules/private/src/index.ts",
      "outside/src/index.ts",
      "modules/private/tsconfig.json",
      "api-reports/README.md",
    ];
    expect(apiRequiredForPaths({ changedPaths: [], publishedPackagePaths: published })).toBe(false);
    for (const file of outside)
      expect(apiRequiredForPaths({ changedPaths: [file], publishedPackagePaths: published })).toBe(
        false,
      );
    for (const file of [
      "package.json",
      "modules/public/package.json",
      "modules/private/package.json",
      "scripts/api-surface-budget.json",
      "scripts/api-surface-budget.ts",
      "scripts/api-reports.ts",
      "tsconfig.base.json",
      "api-reports/new-package/nested/entry.api.md",
      ".github/workflows/ci.yml",
      "scripts/ci-api-plan.ts",
      "scripts/ci-api-plan.test.ts",
      "scripts/ci-run-contract.json",
    ])
      expect(apiRequiredForPaths({ changedPaths: [file], publishedPackagePaths: published })).toBe(
        true,
      );
  });

  test("malformed roots and workspace manifests cannot quietly remove a package from scope", () => {
    for (const workspaces of [
      undefined,
      [],
      "modules/*",
      [3],
      ["../outside/*"],
      ["!modules/*"],
      ["missing/*"],
    ]) {
      const repoRoot = fixture({ workspaces, manifests: { "modules/public": { name: "public" } } });
      expect(() => publishedPackagePaths(repoRoot)).toThrow();
    }
    for (const manifest of [null, [], {}, { name: "" }, { name: "public", private: "true" }]) {
      const repoRoot = fixture({
        workspaces: ["modules/*"],
        manifests: { "modules/public": manifest },
      });
      expect(() => publishedPackagePaths(repoRoot)).toThrow();
    }
    const duplicate = fixture({
      workspaces: ["modules/*"],
      manifests: {
        "modules/one": { name: "same" },
        "modules/two": { name: "same", private: true },
      },
    });
    expect(() => publishedPackagePaths(duplicate)).toThrow();
    write(duplicate, "modules/two/package.json", "{bad JSON");
    expect(() => publishedPackagePaths(duplicate)).toThrow();
  });

  test("valid early matches cannot hide malformed changed paths or package roots", () => {
    for (const invalid of [
      "",
      "..",
      "../outside.ts",
      "/absolute.ts",
      "modules/public/src/../entry.ts",
      "a\0b",
      "a\\b",
    ])
      expect(() =>
        apiRequiredForPaths({
          changedPaths: ["scripts/api-surface-budget.json", invalid],
          publishedPackagePaths: ["modules/public"],
        }),
      ).toThrow();
    expect(() =>
      apiRequiredForPaths({ changedPaths: [], publishedPackagePaths: ["../outside"] }),
    ).toThrow();
  });
});

describe("PR diff and fail-closed CLI", () => {
  test("renames out of published source and deletions retain their source paths", () => {
    const { repoRoot, base } = gitFixture();
    const contents = readFileSync(path.join(repoRoot, "modules/public/src/original.ts"), "utf8");
    rmSync(path.join(repoRoot, "modules/public/src/original.ts"));
    write(repoRoot, "docs/moved.ts", contents);
    const head = snapshot(repoRoot, base);
    const changedPaths = changedApiPaths({ repoRoot, base, head });
    expect(new Set(changedPaths)).toEqual(
      new Set(["docs/moved.ts", "modules/public/src/original.ts"]),
    );
    expect(
      apiRequiredForPaths({ changedPaths, publishedPackagePaths: publishedPackagePaths(repoRoot) }),
    ).toBe(true);
    rmSync(path.join(repoRoot, "docs/moved.ts"));
    const deleted = snapshot(repoRoot, head);
    expect(changedApiPaths({ repoRoot, base, head: deleted })).toEqual([
      "modules/public/src/original.ts",
    ]);
  });

  test("diffs use the merge base and NUL-safe names, including empty changes", () => {
    const { repoRoot, base: ancestor } = gitFixture();
    write(repoRoot, "modules/public/src/base-only.ts", "base branch\n");
    const base = snapshot(repoRoot, ancestor);
    rmSync(path.join(repoRoot, "modules/public/src/base-only.ts"));
    write(repoRoot, "docs/name with\na newline.md", "head branch\n");
    const head = snapshot(repoRoot, ancestor);
    expect(changedApiPaths({ repoRoot, base, head })).toEqual(["docs/name with\na newline.md"]);
    expect(changedApiPaths({ repoRoot, base: head, head })).toEqual([]);
    expect(() => changedApiPaths({ repoRoot, base: "HEAD", head })).toThrow();
    expect(() => changedApiPaths({ repoRoot, base: "0".repeat(40), head })).toThrow();
  });

  test("CLI writes true or false for real diffs, and true plus a failing exit for every boundary failure", () => {
    const { repoRoot, base } = gitFixture();
    write(repoRoot, "docs/guide.md", "updated guide\n");
    write(repoRoot, "modules/private/src/private.ts", "updated private source\n");
    const docsHead = snapshot(repoRoot, base);
    write(repoRoot, "modules/public/src/new.txt", "published source asset\n");
    const sourceHead = snapshot(repoRoot, docsHead);
    const output = path.join(repoRoot, "github-output");
    const run = (args: string[]) => {
      writeFileSync(output, "previous=value\n");
      const result = Bun.spawnSync(
        [process.execPath, path.join(import.meta.dir, "ci-api-plan.ts"), ...args],
        {
          cwd: path.join(repoRoot, "modules/private"),
          stdout: "pipe",
          stderr: "pipe",
          env: { ...process.env, GITHUB_OUTPUT: output },
        },
      );
      return {
        exitCode: result.exitCode,
        output: readFileSync(output, "utf8"),
        stderr: result.stderr.toString(),
      };
    };
    expect(run(["--base", base, "--head", docsHead])).toEqual({
      exitCode: 0,
      output: "previous=value\napi_required=false\n",
      stderr: "",
    });
    expect(run(["--head", sourceHead, "--base", base])).toEqual({
      exitCode: 0,
      output: "previous=value\napi_required=true\n",
      stderr: "",
    });
    for (const args of [
      [],
      ["--base"],
      ["--base", "HEAD", "--head", sourceHead],
      ["--base", "0".repeat(40), "--head", sourceHead],
      ["--base", base, "--base", base, "--head", sourceHead],
      ["--base", base, "--head", sourceHead, "--unknown", "value"],
    ]) {
      const result = run(args);
      expect(result.exitCode).not.toBe(0);
      expect(result.output).toBe("previous=value\napi_required=true\n");
      expect(result.stderr).not.toBe("");
    }
    write(repoRoot, "modules/public/package.json", "{invalid");
    const malformed = run(["--base", base, "--head", docsHead]);
    expect(malformed.exitCode).not.toBe(0);
    expect(malformed.output).toBe("previous=value\napi_required=true\n");
  });
});
