/**
 * Repository source guards that walk files or start processes state a scaled
 * budget, and every test that starts a child process is bounded by what the
 * child does: its file runs under the subprocess budget and no test or child
 * carries a bare wall-clock limit.
 */
import { expect, setDefaultTimeout, test } from "bun:test";
import { existsSync, readFileSync } from "node:fs";
import path from "node:path";

import { runChild, subprocessTestTimeout } from "../test/subprocess";
import {
  type Node,
  parseSourceModule,
  type Program,
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
const SUBPROCESS_BUDGET = "subprocessTestTimeout";
const TEST_FUNCTIONS = new Set(["it", "test"]);
const TIMEOUT_KEYS = new Set(["hangTimeoutMs", "timeout"]);
const UNBOUNDED_SUBPROCESS_TESTS = "scripts/unbounded-subprocess-tests.json";

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

const calleeRoot = (callee: Node): string | undefined => {
  if (callee.type === "Identifier") return callee.name;
  if (callee.type === "MemberExpression") return calleeRoot(callee.object);
  if (callee.type === "CallExpression") return calleeRoot(callee.callee);
  return undefined;
};

const isBudgetCall = (node: Node): boolean =>
  node.type === "CallExpression" && callName(node.callee) === SUBPROCESS_BUDGET;

/** A number written in place: a literal, arithmetic on literals or a constant naming one. */
const isBareNumber = (node: Node, constants: ReadonlySet<string>): boolean => {
  switch (node.type) {
    case "Literal":
      return typeof node.value === "number";
    case "BinaryExpression":
      return (
        isBareNumber(node.left, constants) &&
        isBareNumber(node.right, constants)
      );
    case "Identifier":
      return constants.has(node.name);
    case "TSAsExpression":
    case "TSSatisfiesExpression":
    case "ParenthesizedExpression":
      return isBareNumber(node.expression, constants);
    default:
      return false;
  }
};

/** Top-level `const NAME = <number>` bindings, which a timeout may name. */
const numericConstants = (program: Program): ReadonlySet<string> => {
  const names = new Set<string>();
  for (const statement of program.body) {
    const declaration =
      statement.type === "ExportNamedDeclaration"
        ? statement.declaration
        : statement;
    if (declaration?.type !== "VariableDeclaration") continue;
    for (const declarator of declaration.declarations) {
      if (
        declarator.id.type === "Identifier" &&
        declarator.init !== null &&
        declarator.init !== undefined &&
        isBareNumber(declarator.init, names)
      ) {
        names.add(declarator.id.name);
      }
    }
  }
  return names;
};

/** Why a test file that starts child processes is not bounded by them. */
const subprocessBoundProblems = (file: string, source: string): string[] => {
  const program = parseSourceModule(file, source).program;
  if (!callsAny(program, SUBPROCESS_METHODS)) return [];
  const problems: string[] = [];
  const declaresBudget = program.body.some(
    (statement) =>
      statement.type === "ExpressionStatement" &&
      statement.expression.type === "CallExpression" &&
      callName(statement.expression.callee) === "setDefaultTimeout" &&
      statement.expression.arguments.some(isBudgetCall),
  );
  if (!declaresBudget) {
    problems.push(`no top-level setDefaultTimeout(${SUBPROCESS_BUDGET}())`);
  }
  const constants = numericConstants(program);
  walkAst(program, (node) => {
    if (
      node.type === "Property" &&
      !node.computed &&
      node.key.type === "Identifier" &&
      TIMEOUT_KEYS.has(node.key.name) &&
      isBareNumber(node.value, constants)
    ) {
      problems.push(`bare ${node.key.name} at offset ${String(node.start)}`);
    }
    if (node.type !== "CallExpression") return true;
    const name = callName(node.callee);
    const [first] = node.arguments;
    if (
      name === "setDefaultTimeout" &&
      first !== undefined &&
      isBareNumber(first, constants)
    ) {
      problems.push(`bare setDefaultTimeout at offset ${String(node.start)}`);
    }
    const root = calleeRoot(node.callee);
    const [, body, limit] = node.arguments;
    if (root === undefined || !TEST_FUNCTIONS.has(root) || limit === undefined)
      return true;
    if (isBareNumber(limit, constants)) {
      problems.push(`bare test timeout at offset ${String(limit.start)}`);
    } else if (
      body !== undefined &&
      callsAny(body, SUBPROCESS_METHODS) &&
      !isBudgetCall(limit)
    ) {
      problems.push(
        `test starting a child overrides ${SUBPROCESS_BUDGET} at offset ${String(limit.start)}`,
      );
    }
    return true;
  });
  return problems;
};

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

const unboundedSubprocessTests = (
  sources: ReadonlyMap<string, string>,
): string[] =>
  [...sources]
    .filter(
      ([file, source]) => subprocessBoundProblems(file, source).length > 0,
    )
    .map(([file]) => file)
    .toSorted();

test("tests that start child processes are bounded by the child, not a bare clock", () => {
  const sources = new Map(
    repositoryTestFiles().map((file) => [
      file,
      readSource(path.join(REPO_ROOT, file)) ?? "",
    ]),
  );
  const listed = JSON.parse(
    readSource(path.join(REPO_ROOT, UNBOUNDED_SUBPROCESS_TESTS)) ?? "[]",
  ) as string[];
  // The list only shrinks. A new subprocess test runs its children through
  // test/subprocess.ts under setDefaultTimeout(subprocessTestTimeout()); a
  // listed file that now passes leaves the list.
  const unbounded = unboundedSubprocessTests(sources);
  expect({
    added: unbounded.filter((file) => !listed.includes(file)),
    nowBounded: listed.filter((file) => !unbounded.includes(file)),
  }).toEqual({ added: [], nowBounded: [] });
});

const boundedSource = `import { setDefaultTimeout, test } from "bun:test";
import { runChild, subprocessTestTimeout } from "../test/subprocess";
setDefaultTimeout(subprocessTestTimeout());
test("child", () => runChild(["bun", "--version"]));`;

test("a subprocess test under the shared budget passes", () => {
  expect(
    subprocessBoundProblems("scripts/bounded.test.ts", boundedSource),
  ).toEqual([]);
});

test.each([
  [
    "the runner's default budget",
    'import { test } from "bun:test";\ntest("child", () => Bun.spawnSync(["node", "-e", "0"]));',
  ],
  [
    "a scan budget",
    'import { setDefaultTimeout, test } from "bun:test";\nimport { scanTestTimeout } from "../test/scan-timeout";\nsetDefaultTimeout(scanTestTimeout(5_000));\ntest("child", () => Bun.spawnSync(["node", "-e", "0"]));',
  ],
  [
    "a bare default",
    boundedSource.replace(
      "setDefaultTimeout(subprocessTestTimeout());",
      "setDefaultTimeout(subprocessTestTimeout());\nsetDefaultTimeout(5000);",
    ),
  ],
  [
    "a bare per-test limit",
    `${boundedSource}\ntest("again", () => runChild(["bun", "--version"]), 5_000);`,
  ],
  [
    "a per-test limit from a named number",
    `const LIMIT_MS = 5 * 1000;\n${boundedSource}\ntest("again", () => runChild(["bun", "--version"]), LIMIT_MS);`,
  ],
  [
    "a smaller helper budget on a test that starts a child",
    `${boundedSource}\ntest("again", () => runChild(["bun", "--version"]), scanTestTimeout(5_000));`,
  ],
  [
    "a bare child timeout",
    `${boundedSource}\ntest("again", () => Bun.spawnSync(["bun", "--version"], { timeout: 5000 }));`,
  ],
  [
    "a bare hang guard",
    `${boundedSource}\ntest("again", () => runChild(["bun", "--version"], { hangTimeoutMs: 5000 }));`,
  ],
])("a subprocess test with %s is rejected", (_case, source) => {
  expect(
    unboundedSubprocessTests(new Map([["scripts/unbounded.test.ts", source]])),
  ).toEqual(["scripts/unbounded.test.ts"]);
});
