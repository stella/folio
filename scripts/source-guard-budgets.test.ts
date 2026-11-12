/**
 * Repository source guards that walk files or start processes state a scaled
 * budget, and every test file is linted by the rule that bounds a test that
 * starts a child process by what the child does
 * (`.oxlint-plugins/folio-subprocess-bounds.ts`, with its fixture test in
 * `scripts/subprocess-bounds-lint.test.ts`).
 */
import { expect, setDefaultTimeout, test } from "bun:test";
import {
  existsSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

import oxlintConfig from "../oxlint.config.ts";
import {
  expectExited,
  runChild,
  subprocessTestTimeout,
} from "../test/subprocess";
import {
  type Node,
  parseSourceModule,
  staticMemberName,
  walkAst,
} from "./lib/oxc-ast";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const BUDGET_HELPERS = new Set([
  "propertyTestTimeout",
  "scanTestTimeout",
  "subprocessTestTimeout",
]);
const WALK_METHODS = new Set([
  "readDirectory",
  "readdir",
  "readdirSync",
  "scan",
  "scanSync",
]);
// `$` is the shell template tag (`Bun.$` or an imported `$`).
const SUBPROCESS_METHODS = new Set([
  "$",
  "execFile",
  "execFileSync",
  "runChild",
  "spawn",
  "spawnSync",
]);
const SUBPROCESS_RULE = "folio-subprocess-bounds/bounded-child-tests";
const SUBPROCESS_RULE_PLUGIN = "./.oxlint-plugins/folio-subprocess-bounds.ts";
const SUBPROCESS_RULE_MARKER = "folio-subprocess-bounds(bounded-child-tests)";

// The slowest test lists the repository's test files with one child.
setDefaultTimeout(subprocessTestTimeout());

const readSource = (file: string): string | undefined =>
  existsSync(file) ? readFileSync(file, "utf8") : undefined;

const sourceGuardFiles = (): string[] => {
  const manifest = JSON.parse(
    readSource(path.join(REPO_ROOT, "package.json")) ?? "{}",
  ) as {
    scripts?: Record<string, string>;
  };
  return (manifest.scripts?.["test:source-guards"] ?? "")
    .split(/\s+/u)
    .filter((part) => /\.test\.(?:ts|mjs)$/u.test(part));
};

const callName = (expression: Node): string | undefined =>
  staticMemberName(expression) ??
  (expression.type === "Identifier" ? expression.name : undefined);

const callsAny = (root: Node, names: ReadonlySet<string>): boolean => {
  let found = false;
  walkAst(root, (node) => {
    if (found) return false;
    if (
      node.type === "CallExpression" ||
      node.type === "TaggedTemplateExpression"
    ) {
      const name = callName(
        node.type === "CallExpression" ? node.callee : node.tag,
      );
      if (name !== undefined && names.has(name)) {
        found = true;
        return false;
      }
    }
    return true;
  });
  return found;
};

const missingBudgets = (sources: ReadonlyMap<string, string>): string[] =>
  [...sources]
    .filter(([file, source]) => {
      const parsed = parseSourceModule(file, source).program;
      return (
        (callsAny(parsed, WALK_METHODS) ||
          callsAny(parsed, SUBPROCESS_METHODS)) &&
        !callsAny(parsed, BUDGET_HELPERS)
      );
    })
    .map(([file]) => file)
    .toSorted();

test("repository-walking source guards declare a shared scaled budget", () => {
  const sources = new Map(
    sourceGuardFiles().map((file) => [
      file,
      readSource(path.join(REPO_ROOT, file)) ?? "",
    ]),
  );
  expect(missingBudgets(sources)).toEqual([]);
});

test("a repository walk without a budget is rejected", () => {
  const fixture = new Map([
    [
      "scripts/unbudgeted.test.ts",
      'import { test } from "bun:test";\ntest("walk", () => new Bun.Glob("**/*").scanSync());',
    ],
  ]);
  expect(missingBudgets(fixture)).toEqual(["scripts/unbudgeted.test.ts"]);
});

test("a shell subprocess without a budget is rejected", () => {
  const fixture = new Map([
    [
      "scripts/unbudgeted-shell.test.ts",
      'import { test } from "bun:test";\ntest("shell", () => Bun.$`bun scripts/check.ts`.quiet());',
    ],
  ]);
  expect(missingBudgets(fixture)).toEqual(["scripts/unbudgeted-shell.test.ts"]);
});

const repositoryTestFiles = (): string[] => {
  const listing = runChild(
    [
      "git",
      "ls-files",
      "-z",
      "--cached",
      "--others",
      "--exclude-standard",
      "--",
      "*.test.ts",
      "*.test.tsx",
      "*.test.mts",
      "*.test.mjs",
    ],
    { cwd: REPO_ROOT },
  );
  if (listing.kind !== "exited" || listing.exitCode !== 0) {
    throw new Error(`git ls-files failed: ${JSON.stringify(listing)}`);
  }
  return listing.stdout
    .split("\0")
    .filter((file) => file !== "" && existsSync(path.join(REPO_ROOT, file)))
    .toSorted();
};

type LintOverride = {
  readonly files?: readonly string[];
  readonly rules?: Readonly<Record<string, unknown>>;
};

/** Why the repository lint config does not apply the subprocess bounds rule to these files. */
const subprocessRuleGaps = (
  config: {
    readonly jsPlugins?: readonly string[];
    readonly ignorePatterns?: readonly string[];
    readonly overrides?: readonly LintOverride[];
  },
  testFiles: readonly string[],
): string[] => {
  const gaps: string[] = [];
  if (!(config.jsPlugins ?? []).includes(SUBPROCESS_RULE_PLUGIN)) {
    gaps.push(`${SUBPROCESS_RULE_PLUGIN} is not a lint plugin`);
  }
  const overrides = config.overrides ?? [];
  const globsWith = (levels: readonly string[]) =>
    overrides
      .filter((override) =>
        levels.includes(String(override.rules?.[SUBPROCESS_RULE])),
      )
      .flatMap((override) => override.files ?? [])
      .map((pattern) => new Bun.Glob(pattern));
  const enabled = globsWith(["error"]);
  const weakened = globsWith(["off", "warn"]);
  const ignored = (config.ignorePatterns ?? []).map(
    (pattern) => new Bun.Glob(pattern),
  );
  for (const file of testFiles) {
    if (ignored.some((glob) => glob.match(file))) {
      gaps.push(`${file} is ignored by the lint config`);
    } else if (!enabled.some((glob) => glob.match(file))) {
      gaps.push(`${file} is not covered by ${SUBPROCESS_RULE}`);
    } else if (weakened.some((glob) => glob.match(file))) {
      gaps.push(`${file} weakens ${SUBPROCESS_RULE}`);
    }
  }
  return gaps;
};

/** The paths `bun run lint` hands to oxlint. */
const lintRoots = (): string[] => {
  const manifest = JSON.parse(
    readSource(path.join(REPO_ROOT, "package.json")) ?? "{}",
  ) as { scripts?: Record<string, string> };
  const [, roots = ""] =
    /oxlint -c oxlint\.config\.ts (.*)$/u.exec(
      manifest.scripts?.["lint"] ?? "",
    ) ?? [];
  return roots.split(/\s+/u).filter((root) => root !== "");
};

/** Repository test files that `bun run lint` reaches, and the rest. */
const testFilesByLintReach = (): {
  linted: string[];
  unlinted: string[];
} => {
  const roots = lintRoots();
  const files = repositoryTestFiles();
  const reached = (file: string) =>
    roots.some((root) => file.startsWith(`${root}/`));
  return {
    linted: files.filter(reached),
    unlinted: files.filter((file) => !reached(file)),
  };
};

test("every test file bun run lint reaches is linted by the subprocess bounds rule", () => {
  const { linted } = testFilesByLintReach();
  expect(linted.length).toBeGreaterThan(0);
  expect(subprocessRuleGaps(oxlintConfig, linted)).toEqual([]);
});

// The lint roots leave out the editor extension and the browser suites, whose
// sources predate the shared rule set. Their test files still answer to the
// same rule, run on its own.
test(
  "test files outside the lint roots pass the subprocess bounds rule",
  () => {
    const { unlinted } = testFilesByLintReach();
    expect(unlinted.length).toBeGreaterThan(0);
    const directory = mkdtempSync(
      path.join(tmpdir(), "folio-subprocess-rule-"),
    );
    try {
      const config = path.join(directory, "oxlint.json");
      writeFileSync(
        config,
        JSON.stringify({
          categories: { correctness: "off" },
          jsPlugins: [path.join(REPO_ROOT, SUBPROCESS_RULE_PLUGIN)],
          rules: { [SUBPROCESS_RULE]: "error" },
        }),
      );
      const lint = expectExited(
        runChild(
          ["bun", "--bun", "oxlint", "-c", config, "--no-ignore", ...unlinted],
          { cwd: REPO_ROOT },
        ),
      );
      const output = `${lint.stdout}${lint.stderr}`;
      expect(output).not.toContain("Failed to load");
      expect({ exitCode: lint.exitCode, output }).toEqual({
        exitCode: 0,
        output: expect.not.stringContaining(SUBPROCESS_RULE_MARKER),
      });
    } finally {
      rmSync(directory, { recursive: true, force: true });
    }
  },
  subprocessTestTimeout(2),
);

test("a test file the lint config does not reach is reported", () => {
  const config = {
    jsPlugins: [SUBPROCESS_RULE_PLUGIN],
    ignorePatterns: ["**/*.gen.ts"],
    overrides: [
      { files: ["**/*.test.ts"], rules: { [SUBPROCESS_RULE]: "error" } },
      { files: ["scripts/off.test.ts"], rules: { [SUBPROCESS_RULE]: "off" } },
    ],
  };
  expect(
    subprocessRuleGaps(config, [
      "scripts/linted.test.ts",
      "scripts/child.test.mjs",
      "scripts/off.test.ts",
      "scripts/types.gen.ts",
    ]),
  ).toEqual([
    `scripts/child.test.mjs is not covered by ${SUBPROCESS_RULE}`,
    `scripts/off.test.ts weakens ${SUBPROCESS_RULE}`,
    "scripts/types.gen.ts is ignored by the lint config",
  ]);
  expect(subprocessRuleGaps({}, [])).toEqual([
    `${SUBPROCESS_RULE_PLUGIN} is not a lint plugin`,
  ]);
});
