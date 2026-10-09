import { afterEach, describe, expect, test } from "bun:test";
import { existsSync, mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { createRequire } from "node:module";
import { tmpdir } from "node:os";
import path from "node:path";
import {
  checkBrowserConfigDiscovery,
  checkBrowserDiscovery,
  trackedBrowserConfigs,
} from "./ci-browser-discovery";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const requireFromRepo = createRequire(path.join(REPO_ROOT, "package.json"));
const fixtures = new Set<string>();
afterEach(() => {
  for (const directory of fixtures) rmSync(directory, { recursive: true, force: true });
  fixtures.clear();
});
const write = (repoRoot: string, file: string, contents = "") => {
  const destination = path.join(repoRoot, file);
  mkdirSync(path.dirname(destination), { recursive: true });
  writeFileSync(destination, contents);
};
type GitOptions = { repoRoot: string; args: string[] };
const git = ({ repoRoot, args }: GitOptions) => {
  const result = Bun.spawnSync(["git", ...args], { cwd: repoRoot, stdout: "pipe", stderr: "pipe" });
  if (result.exitCode !== 0) throw new TypeError(result.stderr.toString());
  return result.stdout.toString();
};
const fixture = () => {
  const repoRoot = mkdtempSync(path.join(tmpdir(), "folio-browser-discovery-"));
  fixtures.add(repoRoot);
  git({ repoRoot, args: ["init", "--quiet"] });
  return repoRoot;
};
const report = JSON.stringify({
  errors: [],
  suites: [
    {
      title: "suite",
      specs: [],
      suites: [
        {
          title: "nested",
          specs: [{ title: "case", tests: [{ projectName: "first" }, { projectName: "second" }] }],
        },
      ],
    },
  ],
});
const success = (stdout: string) => ({ exitCode: 0, stdout, stderr: "" });
type Run = NonNullable<Parameters<typeof checkBrowserDiscovery>[0]>["run"];

describe("tracked Playwright config discovery", () => {
  test("all supported extensions, root and nested configs come from Git, including new configs", () => {
    const repoRoot = fixture();
    const configs = [
      "playwright.config.ts",
      "playwright.smoke.config.js",
      "packages/editor/playwright.config.mts",
      "nested/playwright.extra.config.cts",
      "other/playwright.config.mjs",
      "deep/path/playwright.config.cjs",
    ];
    for (const file of configs) write(repoRoot, file);
    for (const file of [
      "playwright.config.json",
      "playwright.config.ts.backup",
      "not-playwright.config.ts",
      "docs/guide.md",
    ])
      write(repoRoot, file);
    git({ repoRoot, args: ["add", "."] });
    write(repoRoot, "untracked/playwright.config.ts");
    expect(trackedBrowserConfigs({ repoRoot })).toEqual(configs.toSorted());
    const added = "packages/new-adapter/playwright.regression.config.ts";
    write(repoRoot, added);
    git({ repoRoot, args: ["add", "--", added] });
    expect(trackedBrowserConfigs({ repoRoot })).toEqual([...configs, added].toSorted());
  });

  test("the current repository's entire tracked config set is enumerated", () => {
    const files = git({ repoRoot: REPO_ROOT, args: ["ls-files", "-z"] }).split("\0");
    const expected = files
      .filter((file) => /(?:^|\/)playwright[^/]*\.config\.(?:ts|js|mts|cts|mjs|cjs)$/u.test(file))
      .toSorted();
    expect(expected.length).toBeGreaterThan(0);
    expect(trackedBrowserConfigs()).toEqual(expected);
  });

  test("Git errors, malformed transport, invalid paths and an empty config set fail closed", () => {
    for (const result of [
      { exitCode: 1, stdout: "", stderr: "not a repository" },
      success(""),
      success("playwright.config.ts"),
      success("docs/guide.md\0"),
      success("../playwright.config.ts\0"),
      success("/playwright.config.ts\0"),
      success("\0"),
    ])
      expect(() => trackedBrowserConfigs({ run: () => result })).toThrow();
  });
});

describe("no-browser discovery preflight", () => {
  test("coverage runs first, then every config uses Node, JSON and --list without project filtering", () => {
    const calls: { command: string[]; cwd: string }[] = [];
    const logs: string[] = [];
    const run: Run = (options) => {
      calls.push(options);
      if (options.command[0] === "git")
        return success("playwright.config.ts\0packages/editor/playwright.config.ts\0");
      if (options.command[0] === "node") return success(report);
      return success("Browser input coverage: 12 kinds registered in Playwright.\n");
    };
    const repoRoot = "/fixture/repository";
    expect(
      checkBrowserDiscovery({
        repoRoot,
        run,
        cliPath: "/installed/playwright/cli.js",
        log: (line) => logs.push(line),
      }),
    ).toEqual([
      { config: "packages/editor/playwright.config.ts", tests: 2 },
      { config: "playwright.config.ts", tests: 2 },
    ]);
    expect(calls.at(0)).toEqual({
      command: [process.execPath, path.join(repoRoot, "scripts/check-browser-input-coverage.ts")],
      cwd: repoRoot,
    });
    expect(calls.at(1)).toEqual({ command: ["git", "ls-files", "-z"], cwd: repoRoot });
    for (const [index, config] of [
      "packages/editor/playwright.config.ts",
      "playwright.config.ts",
    ].entries())
      expect(calls.at(index + 2)).toEqual({
        command: [
          "node",
          "/installed/playwright/cli.js",
          "test",
          "--config",
          path.join(repoRoot, config),
          "--list",
          "--reporter=json",
        ],
        cwd: repoRoot,
      });
    expect(logs).toContain("Browser discovery: 2 tracked configs validated.");
    expect(logs.filter((line) => line.includes("2 tests across all projects.")).length).toBe(2);
  });

  test("every discovery failure stops before subsequent configs, even reporter errors with exit zero", () => {
    const failures = [
      { exitCode: 1, stdout: report, stderr: "module import failed" },
      success(JSON.stringify({ errors: [{ message: "duplicate test title" }], suites: [] })),
      success("not JSON"),
      success("{}"),
      success(JSON.stringify({ errors: [], suites: [] })),
      success(JSON.stringify({ errors: [], suites: [{ specs: {} }] })),
      success(JSON.stringify({ errors: [], suites: [{ specs: [{ tests: null }] }] })),
      success(JSON.stringify({ errors: [], suites: [{ specs: [], suites: {} }] })),
    ];
    for (const failure of failures) {
      const listed: string[][] = [];
      const run: Run = ({ command }) => {
        if (command[0] === "git")
          return success("playwright.config.ts\0playwright.smoke.config.ts\0");
        if (command[0] !== "node") return success("");
        listed.push(command);
        return failure;
      };
      expect(() =>
        checkBrowserDiscovery({ run, cliPath: "/installed/cli.js", log: () => {} }),
      ).toThrow();
      expect(listed.length).toBe(1);
    }
    let calls = 0;
    expect(() =>
      checkBrowserDiscovery({
        run: () => {
          calls += 1;
          return { exitCode: 1, stdout: "", stderr: "missing input kind" };
        },
      }),
    ).toThrow("Browser input coverage failed");
    expect(calls).toBe(1);
  });

  test("the actual pinned Node CLI lists all projects without executing tests or web servers", () => {
    const repoRoot = fixture();
    write(
      repoRoot,
      "scripts/check-browser-input-coverage.ts",
      'process.stdout.write("Fixture coverage registered.\\n");\n',
    );
    const playwrightModule = requireFromRepo.resolve("@playwright/test");
    write(
      repoRoot,
      "playwright.config.cjs",
      `module.exports = { testDir: './tests', projects: [{name:'first'}, {name:'second'}], webServer: {command: 'node -e "process.exit(99)"', port: 61999} };\n`,
    );
    write(
      repoRoot,
      "tests/discovery.spec.cjs",
      `const {test} = require(${JSON.stringify(playwrightModule)});\ntest('discover only', () => { throw new Error('test body must never execute'); });\n`,
    );
    git({ repoRoot, args: ["add", "."] });
    const options = {
      repoRoot,
      cliPath: requireFromRepo.resolve("@playwright/test/cli"),
      log: () => {},
    };
    expect(checkBrowserDiscovery(options)).toEqual([{ config: "playwright.config.cjs", tests: 2 }]);
    expect(existsSync(path.join(repoRoot, "test-results"))).toBe(false);
    const added = "nested/playwright.extra.config.cjs";
    write(
      repoRoot,
      added,
      "module.exports = {testDir:'../tests', projects:[{name:'new-project'}]};\n",
    );
    git({ repoRoot, args: ["add", "--", added] });
    expect(checkBrowserDiscovery(options)).toEqual([
      { config: added, tests: 1 },
      { config: "playwright.config.cjs", tests: 2 },
    ]);
    write(
      repoRoot,
      "tests/discovery.spec.cjs",
      'require("folio-fixture-nonexistent-discovery-module");\n',
    );
    expect(() => checkBrowserDiscovery(options)).toThrow(
      "Playwright discovery failed for nested/playwright.extra.config.cjs",
    );
  }, 20_000);

  test("the repository's browser input coverage passes before discovery", () => {
    const result = Bun.spawnSync(
      [process.execPath, path.join(REPO_ROOT, "scripts/check-browser-input-coverage.ts")],
      { cwd: REPO_ROOT, stdout: "pipe", stderr: "pipe" },
    );
    expect(result.stderr.toString()).toBe("");
    expect(result.exitCode).toBe(0);
  }, 20_000);

  // Each independent Node process gets the same deadline. Sharing a deadline
  // made later graphs inherit the cost of every preceding graph and coverage.
  for (const config of trackedBrowserConfigs()) {
    test(`the repository's real browser import graph loads in Node: ${config}`, () => {
      const result = checkBrowserConfigDiscovery({ config });
      expect(result.config).toBe(config);
      expect(result.tests).toBeGreaterThan(0);
    }, 20_000);
  }

  test("the command line refuses filtering arguments", () => {
    const result = Bun.spawnSync(
      [
        process.execPath,
        path.join(import.meta.dir, "ci-browser-discovery.ts"),
        "--project=interactions",
      ],
      { stdout: "pipe", stderr: "pipe" },
    );
    expect(result.exitCode).toBe(1);
    expect(result.stderr.toString()).toContain("takes no arguments");
  });
});
