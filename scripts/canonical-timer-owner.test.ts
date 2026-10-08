import { expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import ts from "typescript";

const rootDir = resolve(import.meta.dir, "..");
const probeFile = resolve(rootDir, "tests/visual/canonicalTimerProbe.ts");

const fixtureProblems = (file: string, source: string) => {
  let ownsFixture = false;
  let importsPlainTest = false;
  const parsed = ts.createSourceFile(file, source, ts.ScriptTarget.Latest, true);
  for (const node of parsed.statements) {
    if (!ts.isImportDeclaration(node) || !ts.isStringLiteral(node.moduleSpecifier)) continue;
    if (node.importClause?.isTypeOnly) continue;
    const bindings = node.importClause?.namedBindings;
    const specifier = node.moduleSpecifier.text;
    if (
      specifier === "@playwright/test" &&
      (node.importClause?.name !== undefined || (bindings && ts.isNamespaceImport(bindings)))
    )
      importsPlainTest = true;
    if (!bindings || !ts.isNamedImports(bindings)) continue;
    const importsTest = bindings.elements.some(
      (binding) => !binding.isTypeOnly && (binding.propertyName ?? binding.name).text === "test",
    );
    if (!importsTest) continue;
    const target = resolve(dirname(file), specifier);
    if (target === probeFile || `${target}.ts` === probeFile) ownsFixture = true;
    else importsPlainTest = true;
  }
  const problems: string[] = [];
  if (!ownsFixture) problems.push("timer oracle consumers must use the instrumented test fixture");
  if (importsPlainTest)
    problems.push("timer oracle consumers must not import another test fixture");
  return problems;
};

test("every timer oracle consumer installs the probe before navigation", () => {
  // Build the harness import graph once; propagate dependencies backwards to
  // include indirect consumers without maintaining a parallel call-site list.
  const sources = new Map(
    [...new Bun.Glob("tests/**/*.{ts,tsx}").scanSync({ cwd: rootDir })].map((file) => {
      const absoluteFile = resolve(rootDir, file);
      return [absoluteFile, readFileSync(absoluteFile, "utf8")] as const;
    }),
  );
  const dependencies = new Map(
    [...sources].map(
      ([file, source]) =>
        [
          file,
          ts.preProcessFile(source, true, true).importedFiles.flatMap(({ fileName }) => {
            if (!fileName.startsWith(".")) return [];
            const target = resolve(dirname(file), fileName);
            return [target, `${target}.ts`, `${target}.tsx`, resolve(target, "index.ts")].filter(
              (candidate) => sources.has(candidate),
            );
          }),
        ] as const,
    ),
  );
  const reachesProbe = new Set([probeFile]);
  let changed = true;
  while (changed) {
    changed = false;
    for (const [file, imports] of dependencies) {
      if (reachesProbe.has(file) || !imports.some((dependency) => reachesProbe.has(dependency)))
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
    problems.push(...fixtureProblems(file, source).map((problem) => `${file}: ${problem}`));
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
  const parsed = ts.createSourceFile("canonical.ts", source, ts.ScriptTarget.Latest, true);
  let references = 0;
  const visit = (node: ts.Node) => {
    if (ts.isIdentifier(node) && node.text === "setTimeout" && !ts.isTypeQueryNode(node.parent))
      references++;
    if (ts.isStringLiteral(node) && node.text === "setTimeout") references++;
    ts.forEachChild(node, visit);
  };
  visit(parsed);
  return references;
};

test("canonical controller timers use the tagged owner helper", () => {
  const controllerDir = resolve(rootDir, "packages/core/src/controller");
  const files = [...new Bun.Glob("canonical*.ts").scanSync({ cwd: controllerDir })].filter(
    (file) => !file.endsWith(".test.ts") && file !== "canonicalInputTimer.ts",
  );
  expect(files.length).toBeGreaterThan(0);
  for (const file of files)
    expect(rawTimerReferences(readFileSync(resolve(controllerDir, file), "utf8"))).toBe(0);
  for (const schedule of [
    "setTimeout(callback, 25)",
    "window.setTimeout(callback, 25)",
    'window["setTimeout"](callback, 25)',
    'import { setTimeout as schedule } from "node:timers"',
    "const { setTimeout: schedule } = globalThis",
  ])
    expect(rawTimerReferences(schedule)).toBeGreaterThan(0);
  expect(rawTimerReferences("let timer: ReturnType<typeof setTimeout>;")).toBe(0);
});
