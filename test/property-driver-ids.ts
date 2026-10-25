import { panic } from "better-result";
import { readdirSync, readFileSync } from "node:fs";
import path from "node:path";
import {
  type Identifier,
  isCallExpression,
  isComputedPropertyName,
  isExportSpecifier,
  isIdentifier,
  isImportClause,
  isImportDeclaration,
  isImportSpecifier,
  isNamedImports,
  isNamespaceImport,
  isNoSubstitutionTemplateLiteral,
  isObjectLiteralExpression,
  isPropertyAccessExpression,
  isPropertyAssignment,
  isShorthandPropertyAssignment,
  isSpreadAssignment,
  isStringLiteral,
  isTypeQueryNode,
  type Node,
  ScriptTarget,
  type SourceFile,
} from "tsrs/unstable/ast";
import { API, type Checker } from "tsrs/unstable/async";

import { parseSourceModule } from "../scripts/lib/oxc-ast";

const DRIVERS = [
  "assertProperty",
  "assertKnownProperty",
  "assertPinnedProperty",
] as const;

const SOURCE_ROOTS = ["packages", "scripts", "test", "tests"];
const GENERATED_DIRECTORIES = new Set([
  "node_modules",
  "dist",
  "target",
  "coverage",
  "test-results",
  "playwright-report",
  "blob-report",
  "engine-parity-out",
  "engine-parity-ref",
]);

/** Discover import candidates without requiring executables on the runner. */
export const propertyDriverFiles = (root: string): string[] => {
  const files: string[] = [];
  const visit = (directory: string): void => {
    for (const entry of readdirSync(path.join(root, directory), {
      withFileTypes: true,
    })) {
      if (entry.name.startsWith(".") || GENERATED_DIRECTORIES.has(entry.name))
        continue;
      const file = path.posix.join(directory, entry.name);
      if (
        directory === "tests/visual/fixtures" &&
        entry.name.startsWith("parity-tmp-")
      )
        continue;
      if (entry.isDirectory()) {
        visit(file);
        continue;
      }
      if (
        !entry.isFile() ||
        !/\.tsx?$/.test(entry.name) ||
        entry.name.endsWith(".typecheck.ts")
      )
        continue;
      if (
        readFileSync(path.join(root, file), "utf8").includes("property-testing")
      )
        files.push(file);
    }
  };
  for (const directory of SOURCE_ROOTS) visit(directory);
  return files.toSorted();
};

/** Whether a file imports a driver by name or the driver module as a namespace. */
const importsPropertyDriver = (text: string, file: string): boolean =>
  parseSourceModule(file, text).program.body.some(
    (statement) =>
      statement.type === "ImportDeclaration" &&
      /\bproperty-testing(?:\.ts)?$/.test(statement.source.value) &&
      statement.specifiers.some(
        (specifier) =>
          specifier.type === "ImportNamespaceSpecifier" ||
          (specifier.type === "ImportSpecifier" &&
            DRIVERS.some(
              (driver) =>
                driver ===
                (specifier.imported.type === "Identifier"
                  ? specifier.imported.name
                  : specifier.imported.value),
            )),
      ),
  );

const propertyDriverImports = (source: SourceFile) => {
  const imports = new Map<string, string>();
  const namespaces = new Set<string>();
  for (const statement of source.statements) {
    if (
      !isImportDeclaration(statement) ||
      !isStringLiteral(statement.moduleSpecifier) ||
      !/\bproperty-testing(?:\.ts)?$/.test(statement.moduleSpecifier.text)
    )
      continue;
    const bindings = statement.importClause?.namedBindings;
    if (!bindings) continue;
    if (isNamespaceImport(bindings)) {
      namespaces.add(bindings.name.text);
      continue;
    }
    for (const specifier of bindings.elements) {
      const name = specifier.propertyName?.text ?? specifier.name.text;
      if (DRIVERS.some((driver) => driver === name))
        imports.set(specifier.name.text, name);
    }
  }
  return { imports, namespaces };
};

type PropertyDriverSource = {
  source: SourceFile;
  checker: Checker;
  /** Imported driver bindings by symbol id. */
  importedSymbols: Map<number, string>;
  namespaceSymbols: Set<number>;
};

let sharedApi: API | undefined;

type DriverInput = { text: string; file: string };

/** Bind each file on its own (no library, no module resolution), in one snapshot. */
const withPropertyDriverSources = async <T>(
  inputs: readonly DriverInput[],
  read: (sources: PropertyDriverSource[]) => Promise<T>,
): Promise<T> => {
  sharedApi ??= new API({ cwd: process.cwd() });
  const absoluteFiles = inputs.map(({ file }) => path.resolve(file));
  const snapshot = await sharedApi.createSnapshot({
    fileSystem: {
      kind: "layer",
      files: Object.fromEntries(
        inputs.map(({ text }, index) => [absoluteFiles[index], text]),
      ),
    },
    createPrograms: absoluteFiles.map((absoluteFile) => ({
      rootFiles: [absoluteFile],
      compilerOptions: {
        noResolve: true,
        noLib: true,
        target: ScriptTarget.Latest,
      },
    })),
  });
  try {
    const sources: PropertyDriverSource[] = [];
    for (const [
      index,
      program,
    ] of snapshot.operation.createdPrograms.entries()) {
      const absoluteFile = absoluteFiles[index] ?? panic("Unpaired program.");
      const { checker } =
        snapshot.getProject(program.id) ??
        panic(`Cannot bind ${absoluteFile}.`);
      const source =
        (await program.getSourceFile(absoluteFile)) ??
        panic(`Cannot read ${absoluteFile}.`);
      sources.push(await bindPropertyDrivers(source, checker));
    }
    return await read(sources);
  } finally {
    await snapshot.dispose();
  }
};

const bindPropertyDrivers = async (
  source: SourceFile,
  checker: Checker,
): Promise<PropertyDriverSource> => {
  const { imports, namespaces } = propertyDriverImports(source);
  const importedSymbols = new Map<number, string>();
  const namespaceSymbols = new Set<number>();
  for (const statement of source.statements) {
    if (!isImportDeclaration(statement)) continue;
    const bindings = statement.importClause?.namedBindings;
    if (!bindings) continue;
    if (isNamespaceImport(bindings)) {
      const symbol = await checker.getSymbolAtLocation(bindings.name);
      if (symbol && namespaces.has(bindings.name.text))
        namespaceSymbols.add(symbol.id);
      continue;
    }
    if (isNamedImports(bindings)) {
      for (const specifier of bindings.elements) {
        const symbol = await checker.getSymbolAtLocation(specifier.name);
        const driver = imports.get(specifier.name.text);
        if (symbol && driver) importedSymbols.set(symbol.id, driver);
      }
    }
  }
  return { source, checker, importedSymbols, namespaceSymbols };
};

const descendants = (root: Node): Node[] => {
  const nodes: Node[] = [];
  const visit = (node: Node): void => {
    nodes.push(node);
    node.forEachChild(visit);
  };
  visit(root);
  return nodes;
};

/** Inspect imported assertion calls, including aliases, without matching source-string fixtures. */
const callsFromSource = async ({
  source,
  checker,
  importedSymbols,
  namespaceSymbols,
}: PropertyDriverSource) => {
  const calls: {
    driver: string;
    id: string | undefined;
    line: number;
    argumentCount: number;
  }[] = [];
  const candidates = descendants(source).flatMap((node) => {
    if (!isCallExpression(node)) return [];
    const callee = node.expression;
    if (isIdentifier(callee)) return [{ node, binding: callee }];
    if (
      isPropertyAccessExpression(callee) &&
      isIdentifier(callee.expression) &&
      DRIVERS.some((name) => name === callee.name.text)
    )
      return [{ node, binding: callee.expression, member: callee.name.text }];
    return [];
  });
  // One checker request for every candidate binding.
  const symbols =
    candidates.length === 0
      ? []
      : await checker.getSymbolAtLocation(
          candidates.map(({ binding }) => binding),
        );
  for (const [candidate, { node, member }] of candidates.entries()) {
    const symbol = symbols[candidate];
    let driver: string | undefined;
    if (member === undefined) {
      driver = symbol ? importedSymbols.get(symbol.id) : undefined;
    } else if (symbol && namespaceSymbols.has(symbol.id)) {
      driver = member;
    }
    if (driver) {
      const options = node.arguments.at(1);
      let id: string | undefined;
      if (options && isObjectLiteralExpression(options)) {
        const fields = options.properties.flatMap((property, index) => {
          if (isSpreadAssignment(property)) return [];
          const name = property.name;
          if (
            !(isIdentifier(name) || isStringLiteral(name)) ||
            name.text !== "id"
          )
            return [];
          return [{ property, index }];
        });
        const field = fields.length === 1 ? fields.at(0) : undefined;
        if (field && isPropertyAssignment(field.property)) {
          const value = field.property.initializer;
          const canBeOverridden = options.properties
            .slice(field.index + 1)
            .some(
              (property) =>
                isSpreadAssignment(property) ||
                isComputedPropertyName(property.name),
            );
          if (
            !canBeOverridden &&
            (isStringLiteral(value) || isNoSubstitutionTemplateLiteral(value))
          )
            id = value.text;
        }
      }
      calls.push({
        driver,
        id,
        line:
          source.getLineAndCharacterOfPosition(node.getStart(source)).line + 1,
        argumentCount: node.arguments.length,
      });
    }
  }
  return calls;
};

/** Reject imported driver values that escape a direct call, with lexical shadowing respected. */
const referenceProblemsFromSource = async ({
  source,
  checker,
  importedSymbols,
  namespaceSymbols,
}: PropertyDriverSource): Promise<string[]> => {
  if (importedSymbols.size === 0 && namespaceSymbols.size === 0) return [];

  const problems: string[] = [];
  const isTypeOnlyReference = (node: Identifier): boolean => {
    for (
      let parent: Node | undefined = node.parent;
      parent;
      parent = parent.parent
    ) {
      if (isTypeQueryNode(parent)) return true;
    }
    return false;
  };
  const isDirectCall = (node: Identifier): boolean =>
    isCallExpression(node.parent) && node.parent.expression === node;
  const report = (node: Identifier, name: string): void => {
    const line =
      source.getLineAndCharacterOfPosition(node.getStart(source)).line + 1;
    problems.push(
      `${name}: imported property driver value escapes a direct call at line ${line}`,
    );
  };
  const identifiers = descendants(source).filter(
    (node): node is Identifier =>
      isIdentifier(node) && !isTypeOnlyReference(node),
  );
  // One checker request for every identifier.
  const symbols =
    identifiers.length === 0
      ? []
      : await checker.getSymbolAtLocation(identifiers);
  for (const [index, node] of identifiers.entries()) {
    let symbol = symbols[index];
    if (isShorthandPropertyAssignment(node.parent)) {
      symbol = await checker.getShorthandAssignmentValueSymbol(node.parent);
    }
    if (isExportSpecifier(node.parent)) {
      symbol = await checker.getExportSpecifierLocalTargetSymbol(node.parent);
    }
    const driver = symbol ? importedSymbols.get(symbol.id) : undefined;
    if (driver && !isDirectCall(node) && !isImportSpecifier(node.parent)) {
      report(node, driver);
    }
    const namespace = symbol !== undefined && namespaceSymbols.has(symbol.id);
    if (
      namespace &&
      !isImportClause(node.parent) &&
      !isNamespaceImport(node.parent)
    ) {
      const parent = node.parent;
      if (isPropertyAccessExpression(parent) && parent.expression === node) {
        if (
          DRIVERS.some((name) => name === parent.name.text) &&
          !(
            isCallExpression(parent.parent) &&
            parent.parent.expression === parent
          )
        )
          report(node, parent.name.text);
      } else {
        report(node, "namespace");
      }
    }
  }
  return problems;
};

/** A seed key must select exactly one assertion site within its source file. */
export const duplicatePropertyDriverIds = (
  calls: Awaited<ReturnType<typeof propertyDriverIds>>,
) => {
  const firstLines = new Map<string, number>();
  const problems: string[] = [];
  for (const { id, line } of calls) {
    if (id === undefined) continue;
    const first = firstLines.get(id);
    if (first !== undefined) {
      problems.push(
        `${id}: duplicate property ID at lines ${first} and ${line}`,
      );
      continue;
    }
    firstLines.set(id, line);
  }
  return problems;
};

type PropertyDriverAnalysis = {
  calls: Awaited<ReturnType<typeof callsFromSource>>;
  problems: string[];
};

/** Bind each file once for both identity and escaped-reference checks. */
export const propertyDriverAnalyses = async (
  inputs: readonly DriverInput[],
): Promise<PropertyDriverAnalysis[]> => {
  // Utility-only imports cannot contain assertion sites or escaped driver values.
  const bound = inputs.filter(({ text, file }) =>
    importsPropertyDriver(text, file),
  );
  const analyses =
    bound.length === 0
      ? []
      : await withPropertyDriverSources(bound, async (sources) => {
          const results: PropertyDriverAnalysis[] = [];
          for (const source of sources) {
            results.push({
              calls: await callsFromSource(source),
              problems: await referenceProblemsFromSource(source),
            });
          }
          return results;
        });
  const byInput = new Map(
    bound.map((input, index) => [input, analyses[index]]),
  );
  return inputs.map(
    (input) => byInput.get(input) ?? { calls: [], problems: [] },
  );
};

export const propertyDriverAnalysis = async (
  text: string,
  file: string,
): Promise<PropertyDriverAnalysis> =>
  (await propertyDriverAnalyses([{ text, file }])).at(0) ??
  panic("No property driver analysis.");

/** Close the compiler process the analyses share; call once after the last one. */
export const closePropertyDriverAnalysis = async (): Promise<void> => {
  const api = sharedApi;
  sharedApi = undefined;
  await api?.close();
};

export const propertyDriverIds = async (text: string, file: string) =>
  (await propertyDriverAnalysis(text, file)).calls;
export const propertyDriverReferenceProblems = async (
  text: string,
  file: string,
) => (await propertyDriverAnalysis(text, file)).problems;
