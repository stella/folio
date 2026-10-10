import { readFileSync, readdirSync } from "node:fs";
import path from "node:path";
import ts from "typescript";

const DRIVERS = ["assertProperty", "assertKnownProperty", "assertPinnedProperty"] as const;

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
    for (const entry of readdirSync(path.join(root, directory), { withFileTypes: true })) {
      if (entry.name.startsWith(".") || GENERATED_DIRECTORIES.has(entry.name)) continue;
      const file = path.posix.join(directory, entry.name);
      if (directory === "tests/visual/fixtures" && entry.name.startsWith("parity-tmp-")) continue;
      if (entry.isDirectory()) {
        visit(file);
        continue;
      }
      if (!entry.isFile() || !/\.tsx?$/.test(entry.name) || entry.name.endsWith(".typecheck.ts"))
        continue;
      if (readFileSync(path.join(root, file), "utf8").includes("property-testing"))
        files.push(file);
    }
  };
  for (const directory of SOURCE_ROOTS) visit(directory);
  return files.toSorted();
};

const propertyDriverImports = (source: ts.SourceFile) => {
  const imports = new Map<string, string>();
  const namespaces = new Set<string>();
  for (const statement of source.statements) {
    if (
      !ts.isImportDeclaration(statement) ||
      !ts.isStringLiteral(statement.moduleSpecifier) ||
      !/\bproperty-testing(?:\.ts)?$/.test(statement.moduleSpecifier.text)
    )
      continue;
    const bindings = statement.importClause?.namedBindings;
    if (!bindings) continue;
    if (ts.isNamespaceImport(bindings)) {
      namespaces.add(bindings.name.text);
      continue;
    }
    for (const specifier of bindings.elements) {
      const name = specifier.propertyName?.text ?? specifier.name.text;
      if (DRIVERS.some((driver) => driver === name)) imports.set(specifier.name.text, name);
    }
  }
  return { imports, namespaces };
};

const propertyDriverSource = (source: ts.SourceFile) => {
  const { imports, namespaces } = propertyDriverImports(source);
  const options: ts.CompilerOptions = {
    noResolve: true,
    noLib: true,
    target: ts.ScriptTarget.Latest,
  };
  const host = ts.createCompilerHost(options);
  const absoluteFile = ts.sys.resolvePath(source.fileName);
  host.getSourceFile = (name) => (ts.sys.resolvePath(name) === absoluteFile ? source : undefined);
  host.fileExists = (name) => ts.sys.resolvePath(name) === absoluteFile;
  host.readFile = (name) => (ts.sys.resolvePath(name) === absoluteFile ? source.text : undefined);
  const program = ts.createProgram([absoluteFile], options, host);
  const checker = program.getTypeChecker();
  const importedSymbols = new Map<ts.Symbol, string>();
  const namespaceSymbols = new Set<ts.Symbol>();
  for (const statement of source.statements) {
    if (!ts.isImportDeclaration(statement)) continue;
    const bindings = statement.importClause?.namedBindings;
    if (!bindings) continue;
    if (ts.isNamespaceImport(bindings)) {
      const symbol = checker.getSymbolAtLocation(bindings.name);
      if (symbol && namespaces.has(bindings.name.text)) namespaceSymbols.add(symbol);
      continue;
    }
    if (ts.isNamedImports(bindings)) {
      for (const specifier of bindings.elements) {
        const symbol = checker.getSymbolAtLocation(specifier.name);
        const driver = imports.get(specifier.name.text);
        if (symbol && driver) importedSymbols.set(symbol, driver);
      }
    }
  }

  return { source, checker, importedSymbols, namespaceSymbols };
};

/** Inspect imported assertion calls, including aliases, without matching source-string fixtures. */
const callsFromSource = ({
  source,
  checker,
  importedSymbols,
  namespaceSymbols,
}: ReturnType<typeof propertyDriverSource>) => {
  const calls: { driver: string; id: string | undefined; line: number; argumentCount: number }[] =
    [];
  const visit = (node: ts.Node): void => {
    if (ts.isCallExpression(node)) {
      const callee = node.expression;
      let driver: string | undefined;
      if (ts.isIdentifier(callee)) {
        const symbol = checker.getSymbolAtLocation(callee);
        driver = symbol ? importedSymbols.get(symbol) : undefined;
      } else if (
        ts.isPropertyAccessExpression(callee) &&
        ts.isIdentifier(callee.expression) &&
        DRIVERS.some((name) => name === callee.name.text)
      ) {
        const symbol = checker.getSymbolAtLocation(callee.expression);
        if (symbol && namespaceSymbols.has(symbol)) driver = callee.name.text;
      }
      if (driver) {
        const options = node.arguments.at(1);
        let id: string | undefined;
        if (options && ts.isObjectLiteralExpression(options)) {
          const fields = options.properties.flatMap((property, index) => {
            if (ts.isSpreadAssignment(property)) return [];
            const name = property.name;
            if (!(ts.isIdentifier(name) || ts.isStringLiteral(name)) || name.text !== "id")
              return [];
            return [{ property, index }];
          });
          const field = fields.length === 1 ? fields.at(0) : undefined;
          if (field && ts.isPropertyAssignment(field.property)) {
            const value = field.property.initializer;
            const canBeOverridden = options.properties
              .slice(field.index + 1)
              .some(
                (property) =>
                  ts.isSpreadAssignment(property) || ts.isComputedPropertyName(property.name),
              );
            if (
              !canBeOverridden &&
              (ts.isStringLiteral(value) || ts.isNoSubstitutionTemplateLiteral(value))
            )
              id = value.text;
          }
        }
        calls.push({
          driver,
          id,
          line: source.getLineAndCharacterOfPosition(node.getStart(source)).line + 1,
          argumentCount: node.arguments.length,
        });
      }
    }
    ts.forEachChild(node, visit);
  };
  visit(source);
  return calls;
};

/** Reject imported driver values that escape a direct call, with lexical shadowing respected. */
const referenceProblemsFromSource = ({
  source,
  checker,
  importedSymbols,
  namespaceSymbols,
}: ReturnType<typeof propertyDriverSource>): string[] => {
  if (importedSymbols.size === 0 && namespaceSymbols.size === 0) return [];

  const problems: string[] = [];
  const isTypeOnlyReference = (node: ts.Identifier): boolean => {
    for (let parent = node.parent; parent; parent = parent.parent) {
      if (ts.isTypeQueryNode(parent)) return true;
    }
    return false;
  };
  const isDirectCall = (node: ts.Identifier): boolean =>
    ts.isCallExpression(node.parent) && node.parent.expression === node;
  const report = (node: ts.Identifier, name: string): void => {
    const line = source.getLineAndCharacterOfPosition(node.getStart(source)).line + 1;
    problems.push(`${name}: imported property driver value escapes a direct call at line ${line}`);
  };
  const visit = (node: ts.Node): void => {
    if (ts.isIdentifier(node) && !isTypeOnlyReference(node)) {
      let symbol = checker.getSymbolAtLocation(node);
      if (ts.isShorthandPropertyAssignment(node.parent)) {
        symbol = checker.getShorthandAssignmentValueSymbol(node.parent);
      }
      if (ts.isExportSpecifier(node.parent)) {
        symbol = checker.getExportSpecifierLocalTargetSymbol(node.parent);
      }
      const driver = symbol ? importedSymbols.get(symbol) : undefined;
      if (driver && !isDirectCall(node) && !ts.isImportSpecifier(node.parent)) {
        report(node, driver);
      }
      const namespace = symbol !== undefined && namespaceSymbols.has(symbol);
      if (namespace && !ts.isImportClause(node.parent) && !ts.isNamespaceImport(node.parent)) {
        const parent = node.parent;
        if (ts.isPropertyAccessExpression(parent) && parent.expression === node) {
          if (
            DRIVERS.some((name) => name === parent.name.text) &&
            !(ts.isCallExpression(parent.parent) && parent.parent.expression === parent)
          )
            report(node, parent.name.text);
        } else {
          report(node, "namespace");
        }
      }
    }
    ts.forEachChild(node, visit);
  };
  visit(source);
  return problems;
};

/** A seed key must select exactly one assertion site within its source file. */
export const duplicatePropertyDriverIds = (calls: ReturnType<typeof propertyDriverIds>) => {
  const firstLines = new Map<string, number>();
  const problems: string[] = [];
  for (const { id, line } of calls) {
    if (id === undefined) continue;
    const first = firstLines.get(id);
    if (first !== undefined) {
      problems.push(`${id}: duplicate property ID at lines ${first} and ${line}`);
      continue;
    }
    firstLines.set(id, line);
  }
  return problems;
};

/** Bind each file once for both identity and escaped-reference checks. */
export const propertyDriverAnalysis = (text: string, file: string) => {
  const parsed = ts.createSourceFile(file, text, ts.ScriptTarget.Latest, true);
  const { imports, namespaces } = propertyDriverImports(parsed);
  // Utility-only imports cannot contain assertion sites or escaped driver values.
  if (imports.size === 0 && namespaces.size === 0) return { calls: [], problems: [] };
  const source = propertyDriverSource(parsed);
  return { calls: callsFromSource(source), problems: referenceProblemsFromSource(source) };
};

export const propertyDriverIds = (text: string, file: string) =>
  propertyDriverAnalysis(text, file).calls;
export const propertyDriverReferenceProblems = (text: string, file: string) =>
  propertyDriverAnalysis(text, file).problems;
