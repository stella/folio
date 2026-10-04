import path from "node:path";
import ts from "typescript";

export type CanonicalCommandCensusSource = Readonly<{
  file: string;
  source: string;
}>;

type Declaration = Readonly<{
  file: string;
  node: ts.Node;
}>;

type ImportBinding = Readonly<{
  name: string;
  source: string;
}>;

type FileIndex = Readonly<{
  sourceFile: ts.SourceFile;
  declarations: ReadonlyMap<string, Declaration | null>;
  imports: ReadonlyMap<string, ImportBinding | null>;
}>;

type CommandRegistration = Readonly<{
  file: string;
  name: string;
  initializer: ts.Node;
}>;

type ReturnedExpression = Readonly<{
  expression: ts.Node | undefined;
}>;

const canonicalWrapperNames = new Set([
  "withCanonicalCommand",
  "withCanonicalParagraphFormatting",
  "withCanonicalStartParagraphFormatting",
]);
const canonicalCommandsFile = "packages/core/src/prosemirror/canonicalCommands.ts";

const declarationName = (node: ts.Node): string | undefined => {
  if ((ts.isFunctionDeclaration(node) || ts.isClassDeclaration(node)) && node.name !== undefined) {
    return node.name.text;
  }

  if (ts.isVariableDeclaration(node) && ts.isIdentifier(node.name)) return node.name.text;
  return undefined;
};

const moduleSpecifier = (node: ts.Node): string | undefined => {
  if (
    (ts.isImportDeclaration(node) || ts.isExportDeclaration(node)) &&
    node.moduleSpecifier !== undefined &&
    ts.isStringLiteral(node.moduleSpecifier)
  ) {
    return node.moduleSpecifier.text;
  }
  return undefined;
};

const addUnique = <T>(map: Map<string, T | null>, key: string, value: T): void => {
  if (!map.has(key)) {
    map.set(key, value);
    return;
  }
  map.set(key, null);
};

const importedBindings = (sourceFile: ts.SourceFile) => {
  const bindings = new Map<string, ImportBinding | null>();
  for (const statement of sourceFile.statements) {
    const specifier = moduleSpecifier(statement);
    if (!specifier || !ts.isImportDeclaration(statement) || !statement.importClause) continue;

    const clause = statement.importClause;
    if (clause.name) addUnique(bindings, clause.name.text, { name: "default", source: specifier });

    const namedBindings = clause.namedBindings;
    if (namedBindings && ts.isNamespaceImport(namedBindings)) {
      addUnique(bindings, namedBindings.name.text, { name: "*", source: specifier });
    }
    if (namedBindings && ts.isNamedImports(namedBindings)) {
      for (const element of namedBindings.elements) {
        addUnique(bindings, element.name.text, {
          name: (element.propertyName ?? element.name).text,
          source: specifier,
        });
      }
    }
  }
  return bindings;
};

const collectDeclarations = (sourceFile: ts.SourceFile, file: string) => {
  const declarations = new Map<string, Declaration | null>();
  const visit = (node: ts.Node): void => {
    const name = declarationName(node);
    if (name) addUnique(declarations, name, { file, node });
    ts.forEachChild(node, visit);
  };
  visit(sourceFile);
  return declarations;
};

const resolveModuleFile = (
  file: string,
  specifier: string,
  files: ReadonlyMap<string, FileIndex>,
) => {
  if (!specifier.startsWith(".")) return undefined;
  const base = path.posix.normalize(path.posix.join(path.posix.dirname(file), specifier));
  const candidates = [base, `${base}.ts`, `${base}.tsx`, path.posix.join(base, "index.ts")];
  return candidates.find((candidate) => files.has(candidate));
};

const propertyName = (name: ts.PropertyName | undefined): string | undefined => {
  if (!name) return undefined;
  if (ts.isIdentifier(name) || ts.isStringLiteral(name) || ts.isNumericLiteral(name)) {
    return name.text;
  }
  return undefined;
};

const commandInitializer = (command: ts.ObjectLiteralElementLike): ts.Node | undefined => {
  if (ts.isPropertyAssignment(command)) return command.initializer;
  if (ts.isMethodDeclaration(command)) return command;
  if (ts.isShorthandPropertyAssignment(command)) return command.name;
  return undefined;
};

const unwrapExpression = (node: ts.Node): ts.Node => {
  if (
    ts.isParenthesizedExpression(node) ||
    ts.isAsExpression(node) ||
    ts.isTypeAssertionExpression(node) ||
    ts.isSatisfiesExpression(node)
  ) {
    return unwrapExpression(node.expression);
  }
  return node;
};

const commandObjectRegistrations = (sourceFile: ts.SourceFile, file: string) => {
  const registrations: CommandRegistration[] = [];
  const visit = (node: ts.Node): void => {
    if (
      (ts.isShorthandPropertyAssignment(node) ||
        ts.isMethodDeclaration(node) ||
        ts.isGetAccessorDeclaration(node)) &&
      propertyName(node.name) === "commands"
    )
      throw new Error(`${file}: command registry shape cannot be censused`);
    if (ts.isPropertyAssignment(node) && propertyName(node.name) === "commands") {
      const object = unwrapExpression(node.initializer);
      if (!ts.isObjectLiteralExpression(object))
        throw new Error(`${file}: command registry must be an object literal for the census`);
      for (const command of object.properties) {
        const name = propertyName(command.name);
        const initializer = commandInitializer(command);
        if (!name || !initializer)
          throw new Error(
            `${file}: command registration cannot be censused: ${command.getText(sourceFile)}`,
          );
        registrations.push({ file, name, initializer });
      }
    }
    ts.forEachChild(node, visit);
  };
  visit(sourceFile);
  return registrations;
};

const declarationValue = (declaration: Declaration): ts.Node => {
  if (ts.isVariableDeclaration(declaration.node) && declaration.node.initializer) {
    return declaration.node.initializer;
  }
  return declaration.node;
};

const returnedExpressions = (node: ts.Node): ReturnedExpression[] => {
  const callable = unwrapExpression(node);
  if (ts.isArrowFunction(callable) && !ts.isBlock(callable.body)) {
    return [{ expression: callable.body }];
  }

  const body =
    ts.isArrowFunction(callable) ||
    ts.isFunctionExpression(callable) ||
    ts.isFunctionDeclaration(callable) ||
    ts.isMethodDeclaration(callable)
      ? callable.body
      : undefined;
  if (!body) return [];

  const expressions: ReturnedExpression[] = [];
  const visit = (current: ts.Node): void => {
    if (
      current !== body &&
      (ts.isArrowFunction(current) ||
        ts.isFunctionExpression(current) ||
        ts.isFunctionDeclaration(current) ||
        ts.isMethodDeclaration(current) ||
        ts.isClassDeclaration(current) ||
        ts.isClassExpression(current))
    ) {
      return;
    }
    if (ts.isReturnStatement(current)) {
      expressions.push({ expression: current.expression });
      return;
    }
    ts.forEachChild(current, visit);
  };
  visit(body);
  return expressions;
};

/**
 * Return extension command registrations without a structural proof that all
 * factory return paths resolve to a canonical command descriptor wrapper.
 * Inputs are source snapshots, so callers can census working-tree and git-base
 * contents with the same pure function.
 */
export const censusUndescribedCanonicalCommands = (
  sources: readonly CanonicalCommandCensusSource[],
): string[] => {
  const files = new Map<string, FileIndex>();
  for (const { file, source } of sources) {
    const normalizedFile = path.posix.normalize(file);
    const sourceFile = ts.createSourceFile(
      normalizedFile,
      source,
      ts.ScriptTarget.Latest,
      true,
      normalizedFile.endsWith(".tsx") ? ts.ScriptKind.TSX : ts.ScriptKind.TS,
    );
    files.set(normalizedFile, {
      sourceFile,
      declarations: collectDeclarations(sourceFile, normalizedFile),
      imports: importedBindings(sourceFile),
    });
  }

  const resolveDeclaration = (file: string, name: string): Declaration | undefined => {
    const fileIndex = files.get(file);
    if (!fileIndex) return undefined;
    if (fileIndex.imports.has(name)) {
      const imported = fileIndex.imports.get(name);
      if (!imported) return undefined;
      const importedFile = resolveModuleFile(file, imported.source, files);
      return importedFile
        ? (files.get(importedFile)?.declarations.get(imported.name) ?? undefined)
        : undefined;
    }
    return fileIndex.declarations.get(name) ?? undefined;
  };

  const isCanonicalWrapper = (file: string, name: string): boolean => {
    const imported = files.get(file)?.imports.get(name);
    if (!imported || !canonicalWrapperNames.has(imported.name)) return false;
    return resolveModuleFile(file, imported.source, files) === canonicalCommandsFile;
  };

  const proveDeclaration = (declaration: Declaration, visited: Set<string>): boolean => {
    const declarationNameValue = declarationName(declaration.node);
    const key = `${declaration.file}#${declarationNameValue ?? declaration.node.pos}`;
    if (visited.has(key)) return false;
    const nextVisited = new Set(visited);
    nextVisited.add(key);

    const value = declarationValue(declaration);
    if (ts.isVariableDeclaration(declaration.node)) {
      if (declaration.node.initializer === undefined) return false;
      const initializer = unwrapExpression(declaration.node.initializer);
      if (ts.isArrowFunction(initializer) || ts.isFunctionExpression(initializer)) {
        const returns = returnedExpressions(initializer);
        return (
          returns.length > 0 &&
          returns.every(
            ({ expression }) =>
              expression !== undefined &&
              proveExpression(declaration.file, expression, nextVisited),
          )
        );
      }
      return proveExpression(declaration.file, value, nextVisited);
    }

    const returns = returnedExpressions(value);
    return (
      returns.length > 0 &&
      returns.every(
        ({ expression }) =>
          expression !== undefined && proveExpression(declaration.file, expression, nextVisited),
      )
    );
  };

  const proveExpression = (file: string, input: ts.Node, visited: Set<string>): boolean => {
    const expression = unwrapExpression(input);
    if (
      ts.isCallExpression(expression) &&
      ts.isIdentifier(unwrapExpression(expression.expression))
    ) {
      const callee = unwrapExpression(expression.expression);
      if (ts.isIdentifier(callee) && isCanonicalWrapper(file, callee.text)) return true;
      if (ts.isIdentifier(callee)) {
        const declaration = resolveDeclaration(file, callee.text);
        return declaration !== undefined && proveDeclaration(declaration, visited);
      }
      return false;
    }

    if (ts.isIdentifier(expression)) {
      if (isCanonicalWrapper(file, expression.text)) return true;
      const declaration = resolveDeclaration(file, expression.text);
      return declaration !== undefined && proveDeclaration(declaration, visited);
    }

    if (ts.isConditionalExpression(expression)) {
      return (
        proveExpression(file, expression.whenTrue, visited) &&
        proveExpression(file, expression.whenFalse, visited)
      );
    }
    return false;
  };

  const provesFactory = (file: string, initializer: ts.Node): boolean => {
    if (ts.isIdentifier(unwrapExpression(initializer))) {
      return proveExpression(file, initializer, new Set());
    }
    const returns = returnedExpressions(initializer);
    return (
      returns.length > 0 &&
      returns.every(
        ({ expression }) =>
          expression !== undefined && proveExpression(file, expression, new Set()),
      )
    );
  };

  const registrations = sources.flatMap(({ file }) => {
    const normalizedFile = path.posix.normalize(file);
    const sourceFile = files.get(normalizedFile)?.sourceFile;
    return sourceFile ? commandObjectRegistrations(sourceFile, normalizedFile) : [];
  });

  return registrations
    .filter(({ file, initializer }) => !provesFactory(file, initializer))
    .map(({ file, name }) => `${file}#${name}`)
    .sort();
};

/** Compare registration occurrences, so a duplicate cannot increase the count under an old id. */
export const checkCanonicalCommandBaseline = (
  current: readonly string[],
  baseline: readonly string[],
) => {
  const available = new Map<string, number>();
  for (const command of baseline) available.set(command, (available.get(command) ?? 0) + 1);
  const additions: string[] = [];
  for (const command of current) {
    const count = available.get(command) ?? 0;
    if (count === 0) additions.push(command);
    else available.set(command, count - 1);
  }
  return additions;
};
