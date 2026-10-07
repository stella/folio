import { expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import ts from "typescript";
import { canonicalTimerOwner, isCanonicalInputTimer } from "../tests/parity/canonicalTimerOwner";

const captured = (owner: string, caller: string) =>
  `Error\n    at window.setTimeout (<anonymous>:10:39)\n    at ${owner}\n    at ${caller}`;
const root = "http://localhost:4200/@fs/work/folio/packages";
const input = `Object.handleKeyDown (${root}/core/src/controller/canonicalInput.ts:119:33)`;

test("a layout timer called through canonical history belongs to the layout", () => {
  const owner = `${root}/react/src/paged-editor/PagedEditor.tsx:1339:50`;
  const stack = captured(owner, input);
  expect(canonicalTimerOwner(stack)).toBe(`at ${owner}`);
  expect(isCanonicalInputTimer(stack)).toBe(false);
});

test("composition and input timers are recognized by their own frame", () => {
  for (const module of ["canonicalComposition", "canonicalInput"]) {
    expect(
      isCanonicalInputTimer(
        captured(`schedule (${root}/core/src/controller/${module}.ts:98:13)`, input),
      ),
    ).toBe(true);
  }
});

test("an unrelated same-named module does not own canonical input", () => {
  expect(isCanonicalInputTimer(captured(`${root}/react/src/canonicalInput.ts:1:1`, input))).toBe(
    false,
  );
});

test("missing owner capture fails instead of declaring a clean state", () => {
  for (const stack of ["unavailable", "Error\n    at window.setTimeout (<anonymous>:10:39)"])
    expect(() => isCanonicalInputTimer(stack)).toThrow("Timer capture has no owner frame");
});

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
