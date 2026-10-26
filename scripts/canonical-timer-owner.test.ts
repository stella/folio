import { expect, setDefaultTimeout, test } from "bun:test";
import { readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";

import { scanTestTimeout } from "../test/scan-timeout";
import { type Node, parseSourceModule, walkAst } from "./lib/oxc-ast";

setDefaultTimeout(scanTestTimeout(30_000));

const rootDir = resolve(import.meta.dir, "..");
const probeFile = resolve(rootDir, "tests/visual/canonicalTimerProbe.ts");

const fixtureProblems = (file: string, source: string) => {
  let ownsFixture = false;
  let importsPlainTest = false;
  const { program } = parseSourceModule(file, source);
  for (const node of program.body) {
    if (node.type !== "ImportDeclaration") continue;
    if (node.importKind === "type") continue;
    const specifier = node.source.value;
    if (
      specifier === "@playwright/test" &&
      node.specifiers.some(
        (binding) =>
          binding.type === "ImportDefaultSpecifier" ||
          binding.type === "ImportNamespaceSpecifier",
      )
    )
      importsPlainTest = true;
    const importsTest = node.specifiers.some(
      (binding) =>
        binding.type === "ImportSpecifier" &&
        binding.importKind !== "type" &&
        (binding.imported.type === "Identifier"
          ? binding.imported.name
          : binding.imported.value) === "test",
    );
    if (!importsTest) continue;
    const target = resolve(dirname(file), specifier);
    if (target === probeFile || `${target}.ts` === probeFile)
      ownsFixture = true;
    else importsPlainTest = true;
  }
  const problems: string[] = [];
  if (!ownsFixture)
    problems.push(
      "timer oracle consumers must use the instrumented test fixture",
    );
  if (importsPlainTest)
    problems.push(
      "timer oracle consumers must not import another test fixture",
    );
  return problems;
};

const literalText = (node: Node | null | undefined) =>
  node?.type === "Literal" && typeof node.value === "string"
    ? node.value
    : undefined;

/** The module a node names, if it names one. */
const moduleNameOf = (node: Node): string | undefined => {
  switch (node.type) {
    case "ImportDeclaration":
    case "ExportAllDeclaration":
    case "ExportNamedDeclaration":
    case "ImportExpression":
    case "TSImportType":
      return literalText(node.source);
    case "TSExternalModuleReference":
      return literalText(node.expression);
    case "CallExpression":
      return node.callee.type === "Identifier" && node.callee.name === "require"
        ? literalText(node.arguments.at(0))
        : undefined;
    default:
      return undefined;
  }
};

/** Every module a file names: imports, re-exports, `import()`, `require()` and import types. */
const importedFiles = (file: string, source: string) => {
  const { program } = parseSourceModule(file, source);
  const names: string[] = [];
  walkAst(program, (node) => {
    const name = moduleNameOf(node);
    if (name !== undefined) names.push(name);
  });
  return names;
};

test("every timer oracle consumer installs the probe before navigation", () => {
  // Build the harness import graph once; propagate dependencies backwards to
  // include indirect consumers without maintaining a parallel call-site list.
  const sources = new Map(
    [...new Bun.Glob("tests/**/*.{ts,tsx}").scanSync({ cwd: rootDir })].map(
      (file) => {
        const absoluteFile = resolve(rootDir, file);
        return [absoluteFile, readFileSync(absoluteFile, "utf8")] as const;
      },
    ),
  );
  const dependencies = new Map(
    [...sources].map(
      ([file, source]) =>
        [
          file,
          importedFiles(file, source).flatMap((fileName) => {
            if (!fileName.startsWith(".")) return [];
            const target = resolve(dirname(file), fileName);
            return [
              target,
              `${target}.ts`,
              `${target}.tsx`,
              resolve(target, "index.ts"),
            ].filter((candidate) => sources.has(candidate));
          }),
        ] as const,
    ),
  );
  const reachesProbe = new Set([probeFile]);
  let changed = true;
  while (changed) {
    changed = false;
    for (const [file, imports] of dependencies) {
      if (
        reachesProbe.has(file) ||
        !imports.some((dependency) => reachesProbe.has(dependency))
      )
        continue;
      reachesProbe.add(file);
      changed = true;
    }
  }
  const problems: string[] = [];
  let consumers = 0;
  for (const [file, source] of sources) {
    if (!file.endsWith(".spec.ts") || !reachesProbe.has(file)) continue;
    consumers++;
    problems.push(
      ...fixtureProblems(file, source).map((problem) => `${file}: ${problem}`),
    );
    // Every detected consumer exercises the rejection oracle too, including
    // aliases and future imports with a different layout.
    const plainFixture = 'import { test } from "@playwright/test";';
    expect(fixtureProblems(file, plainFixture).length).toBeGreaterThan(0);
    const instrumentedFixture = 'import { test } from "./canonicalTimerProbe";';
    for (const plainImport of [
      'import * as playwright from "@playwright/test";',
      'import playwright from "@playwright/test";',
      'import { test as plain } from "./otherFixture";',
    ]) {
      expect(
        fixtureProblems(file, `${instrumentedFixture}\n${plainImport}`).length,
      ).toBeGreaterThan(0);
    }
  }
  expect(consumers).toBeGreaterThan(0);
  expect(problems).toEqual([]);
});

const rawTimerReferences = (source: string) => {
  const { program } = parseSourceModule("canonical.ts", source);
  let references = 0;
  // oxc repeats a shorthand name as key and value; count each token once.
  const counted = new Set<number>();
  walkAst(program, (node, ancestors) => {
    if (
      node.type === "Identifier" &&
      node.name === "setTimeout" &&
      ancestors.at(-1)?.type !== "TSTypeQuery" &&
      !counted.has(node.start)
    ) {
      counted.add(node.start);
      references++;
    }
    if (node.type === "Literal" && node.value === "setTimeout") references++;
  });
  return references;
};

test("canonical controller timers use the tagged owner helper", () => {
  const controllerDir = resolve(rootDir, "packages/core/src/controller");
  const files = [
    ...new Bun.Glob("canonical*.ts").scanSync({ cwd: controllerDir }),
  ].filter(
    (file) => !file.endsWith(".test.ts") && file !== "canonicalInputTimer.ts",
  );
  expect(files.length).toBeGreaterThan(0);
  for (const file of files)
    expect(
      rawTimerReferences(readFileSync(resolve(controllerDir, file), "utf8")),
    ).toBe(0);
  for (const schedule of [
    "setTimeout(callback, 25)",
    "window.setTimeout(callback, 25)",
    'window["setTimeout"](callback, 25)',
    'import { setTimeout as schedule } from "node:timers"',
    "const { setTimeout: schedule } = globalThis",
  ])
    expect(rawTimerReferences(schedule)).toBeGreaterThan(0);
  expect(rawTimerReferences("let timer: ReturnType<typeof setTimeout>;")).toBe(
    0,
  );
});
