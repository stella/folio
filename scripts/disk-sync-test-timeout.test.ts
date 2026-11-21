import { expect, setDefaultTimeout, test } from "bun:test";
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { dirname, join, relative, resolve } from "node:path";

import { DISK_SYNC_TEST_TIMEOUT_MS } from "../test/disk-sync-timeout";
import { scanTestTimeout } from "../test/scan-timeout";
import {
  moduleReferences,
  type Node,
  parseSourceModule,
  walkAst,
} from "./lib/oxc-ast";

setDefaultTimeout(scanTestTimeout(30_000));

const rootDir = resolve(import.meta.dir, "..");
const cliDir = resolve(rootDir, "packages/cli/src");
const BUDGET = "DISK_SYNC_TEST_TIMEOUT_MS";
/** Budgets at least as long as {@link BUDGET}: a child's hang guard alone exceeds it. */
const LONGER_BUDGETS: ReadonlySet<string> = new Set(["subprocessTestTimeout"]);

const sourceFiles = (dir: string): string[] =>
  readdirSync(dir, { recursive: true, encoding: "utf8" })
    .filter((name) => name.endsWith(".ts"))
    .map((name) => join(dir, name));

const parsed = new Map(
  sourceFiles(cliDir).map((file) => [
    file,
    parseSourceModule(file, readFileSync(file, "utf8")),
  ]),
);

const programOf = (file: string) =>
  parsed.get(file)?.program ??
  (() => {
    throw new Error(`Unparsed module ${relative(rootDir, file)}`);
  })();

/** An `await handle.sync()`: the module flushes a file to disk. */
const syncsToDisk = (file: string): boolean => {
  let found = false;
  walkAst(programOf(file), (node: Node) => {
    if (
      node.type === "AwaitExpression" &&
      node.argument.type === "CallExpression" &&
      node.argument.callee.type === "MemberExpression" &&
      !node.argument.callee.computed &&
      node.argument.callee.property.type === "Identifier" &&
      node.argument.callee.property.name === "sync"
    )
      found = true;
  });
  return found;
};

const resolveLocal = (from: string, specifier: string): string | undefined => {
  if (!specifier.startsWith(".")) return undefined;
  const base = resolve(dirname(from), specifier);
  return [base, `${base}.ts`].find(
    (candidate) => candidate.endsWith(".ts") && existsSync(candidate),
  );
};

const reachesSync = (start: string): boolean => {
  const seen = new Set<string>();
  const pending = [start];
  for (let file = pending.pop(); file !== undefined; file = pending.pop()) {
    if (seen.has(file)) continue;
    seen.add(file);
    if (syncsToDisk(file)) return true;
    for (const { specifier, typeOnly } of moduleReferences(programOf(file))) {
      if (typeOnly) continue;
      const target = resolveLocal(file, specifier);
      if (target !== undefined && parsed.has(target)) pending.push(target);
    }
  }
  return false;
};

/** The callee or identifier a top-level `setDefaultTimeout(...)` passes. */
const declaredBudget = (file: string): string | undefined => {
  for (const statement of programOf(file).body) {
    if (statement.type !== "ExpressionStatement") continue;
    const call = statement.expression;
    if (
      call.type !== "CallExpression" ||
      call.callee.type !== "Identifier" ||
      call.callee.name !== "setDefaultTimeout"
    )
      continue;
    const argument = call.arguments.at(0);
    if (argument?.type === "Identifier") return argument.name;
    if (
      argument?.type === "CallExpression" &&
      argument.callee.type === "Identifier"
    )
      return argument.callee.name;
    return "";
  }
  return undefined;
};

const TEST_FUNCTIONS: ReadonlySet<string> = new Set(["test", "it"]);

/**
 * The function a call's callee chain starts from: `test` for `test(...)`,
 * `test.only(...)`, `test.each(table)(...)` and `test.if(cond)(...)`.
 */
const rootName = (callee: Node): string | undefined => {
  let node: Node = callee;
  for (;;) {
    if (node.type === "Identifier") return node.name;
    if (node.type === "MemberExpression") node = node.object;
    else if (node.type === "CallExpression") node = node.callee;
    else return undefined;
  }
};

/** A literal budget a test call states: `ms` or `{ timeout: ms }`. */
const literalBudget = (argument: Node | undefined): number | undefined => {
  if (argument?.type === "Literal" && typeof argument.value === "number")
    return argument.value;
  if (argument?.type !== "ObjectExpression") return undefined;
  for (const property of argument.properties) {
    if (
      property.type === "Property" &&
      !property.computed &&
      property.key.type === "Identifier" &&
      property.key.name === "timeout" &&
      property.value.type === "Literal" &&
      typeof property.value.value === "number"
    )
      return property.value.value;
  }
  return undefined;
};

/** Per-test literal budgets shorter than the shared one, in any test call form. */
const shorterLiteralBudgets = (file: string): number[] => {
  const shorter: number[] = [];
  walkAst(programOf(file), (node: Node) => {
    if (node.type !== "CallExpression") return;
    if (!TEST_FUNCTIONS.has(rootName(node.callee) ?? "")) return;
    const budget = literalBudget(node.arguments.at(2));
    if (budget !== undefined && budget < DISK_SYNC_TEST_TIMEOUT_MS)
      shorter.push(budget);
  });
  return shorter;
};

const syncingTests = [...parsed.keys()]
  .filter((file) => file.endsWith(".test.ts") && reachesSync(file))
  .toSorted();

test("the CLI modules that sync a write to disk are found", () => {
  expect(
    [...parsed.keys()]
      .filter(syncsToDisk)
      .map((file) => relative(cliDir, file))
      .toSorted(),
  ).toEqual(["journal.ts", "output-file.ts", "sidecar.ts"]);
});

test("every CLI test reaching a disk sync runs under the disk-sync budget", () => {
  expect(syncingTests.length).toBeGreaterThan(0);
  const unbudgeted = syncingTests
    .filter((file) => {
      const budget = declaredBudget(file);
      return budget !== BUDGET && !LONGER_BUDGETS.has(budget ?? "");
    })
    .map((file) => relative(rootDir, file));
  expect(unbudgeted).toEqual([]);
  const shortened = syncingTests
    .filter((file) => shorterLiteralBudgets(file).length > 0)
    .map((file) => relative(rootDir, file));
  expect(shortened).toEqual([]);
});

test("the guard reads a planted unbudgeted syncing test", () => {
  const file = join(cliDir, "planted.test.ts");
  parsed.set(
    file,
    parseSourceModule(
      file,
      [
        'import { test } from "bun:test";',
        'import { saveTo } from "./journal";',
        'test("plain", () => {}, 5_000);',
        'test.each([1])("each", () => {}, 4_000);',
        'test("options", () => {}, { timeout: 3_000 });',
        'test.only("long enough", () => {}, { timeout: 30_000 });',
        'test.if(true)("long enough too", () => {}, 60_000);',
      ].join("\n"),
    ),
  );
  try {
    expect(reachesSync(file)).toBe(true);
    expect(declaredBudget(file)).toBeUndefined();
    expect(shorterLiteralBudgets(file)).toEqual([5_000, 4_000, 3_000]);
  } finally {
    parsed.delete(file);
  }
});
