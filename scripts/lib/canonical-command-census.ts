import path from "node:path";

import { type Node, nodeText, parseSourceModule, walkAst } from "./oxc-ast";

export type CanonicalCommandCensusSource = Readonly<{
  file: string;
  source: string;
}>;

type Declaration = Readonly<{
  file: string;
  node: Node;
}>;

type ImportBinding = Readonly<{
  name: string;
  source: string;
}>;

type FileIndex = Readonly<{
  program: Node;
  text: string;
  declarations: ReadonlyMap<string, Declaration | null>;
  imports: ReadonlyMap<string, ImportBinding | null>;
}>;

type CommandRegistration = Readonly<{
  file: string;
  name: string;
  initializer: Node;
}>;

type ReturnedExpression = Readonly<{
  expression: Node | undefined;
}>;

const canonicalWrapperNames = new Set([
  "withCanonicalCommand",
  "withCanonicalParagraphFormatting",
  "withCanonicalStartParagraphFormatting",
]);
const canonicalCommandsFile =
  "packages/core/src/prosemirror/canonicalCommands.ts";

const declarationName = (node: Node): string | undefined => {
  if (
    (node.type === "FunctionDeclaration" ||
      node.type === "TSDeclareFunction" ||
      node.type === "ClassDeclaration") &&
    node.id !== null
  ) {
    return node.id.name;
  }

  if (node.type === "VariableDeclarator" && node.id.type === "Identifier")
    return node.id.name;
  // A catch binding is a variable declaration without an initializer.
  if (node.type === "CatchClause" && node.param?.type === "Identifier")
    return node.param.name;
  return undefined;
};

const addUnique = <T>(
  map: Map<string, T | null>,
  key: string,
  value: T,
): void => {
  if (!map.has(key)) {
    map.set(key, value);
    return;
  }
  map.set(key, null);
};

const importedBindings = (program: Node) => {
  const bindings = new Map<string, ImportBinding | null>();
  if (program.type !== "Program") return bindings;
  for (const statement of program.body) {
    if (statement.type !== "ImportDeclaration") continue;
    const specifier = statement.source.value;
    if (!specifier) continue;

    for (const binding of statement.specifiers) {
      if (binding.type === "ImportDefaultSpecifier")
        addUnique(bindings, binding.local.name, {
          name: "default",
          source: specifier,
        });
      if (binding.type === "ImportNamespaceSpecifier")
        addUnique(bindings, binding.local.name, {
          name: "*",
          source: specifier,
        });
      if (binding.type === "ImportSpecifier")
        addUnique(bindings, binding.local.name, {
          name:
            binding.imported.type === "Identifier"
              ? binding.imported.name
              : binding.imported.value,
          source: specifier,
        });
    }
  }
  return bindings;
};

const collectDeclarations = (program: Node, file: string) => {
  const declarations = new Map<string, Declaration | null>();
  walkAst(program, (node) => {
    const name = declarationName(node);
    if (name) addUnique(declarations, name, { file, node });
  });
  return declarations;
};

const resolveModuleFile = (
  file: string,
  specifier: string,
  files: ReadonlyMap<string, FileIndex>,
) => {
  if (!specifier.startsWith(".")) return undefined;
  const base = path.posix.normalize(
    path.posix.join(path.posix.dirname(file), specifier),
  );
  const candidates = [
    base,
    `${base}.ts`,
    `${base}.tsx`,
    path.posix.join(base, "index.ts"),
  ];
  return candidates.find((candidate) => files.has(candidate));
};

/** A static identifier, string or numeric member name. */
const propertyName = (member: Node): string | undefined => {
  if (
    member.type !== "Property" &&
    member.type !== "MethodDefinition" &&
    member.type !== "TSAbstractMethodDefinition"
  )
    return undefined;
  if (member.computed) return undefined;
  const { key } = member;
  if (key.type === "Identifier") return key.name;
  if (key.type === "Literal" && typeof key.value === "string") return key.value;
  if (key.type === "Literal" && typeof key.value === "number")
    return String(key.value);
  return undefined;
};

const commandInitializer = (command: Node): Node | undefined => {
  if (command.type !== "Property" || command.kind !== "init") return undefined;
  if (command.shorthand) return command.key;
  return command.value;
};

/** Parentheses, type assertions and optional chains do not change the value. */
const unwrapExpression = (node: Node): Node => {
  if (
    node.type === "ParenthesizedExpression" ||
    node.type === "TSAsExpression" ||
    node.type === "TSTypeAssertion" ||
    node.type === "TSSatisfiesExpression" ||
    node.type === "ChainExpression"
  ) {
    return unwrapExpression(node.expression);
  }
  return node;
};

const isFunctionScope = (node: Node) =>
  node.type === "ArrowFunctionExpression" ||
  node.type === "FunctionExpression" ||
  node.type === "FunctionDeclaration" ||
  node.type === "MethodDefinition" ||
  node.type === "TSAbstractMethodDefinition" ||
  node.type === "ClassDeclaration" ||
  node.type === "ClassExpression";

const returnedExpressions = (node: Node): ReturnedExpression[] => {
  const callable = unwrapExpression(node);
  if (callable.type === "ArrowFunctionExpression" && callable.expression) {
    return [{ expression: callable.body }];
  }

  const body =
    callable.type === "ArrowFunctionExpression" ||
    callable.type === "FunctionExpression" ||
    callable.type === "FunctionDeclaration"
      ? callable.body
      : undefined;
  if (!body) return [];

  const expressions: ReturnedExpression[] = [];
  walkAst(body, (current, ancestors) => {
    if (current === body) return true;
    if (isFunctionScope(current)) {
      // Object accessors are not separate methods; their returns stay visible.
      const parent = ancestors.at(-1);
      return (
        current.type === "FunctionExpression" &&
        parent?.type === "Property" &&
        parent.kind !== "init"
      );
    }
    if (current.type === "ReturnStatement") {
      expressions.push({ expression: current.argument ?? undefined });
      return false;
    }
    return true;
  });
  return expressions;
};

const commandObjectRegistrations = (
  program: Node,
  text: string,
  file: string,
) => {
  const registrations: CommandRegistration[] = [];
  walkAst(program, (node, ancestors) => {
    if (propertyName(node) !== "commands") return;
    if (
      node.type === "MethodDefinition" ||
      node.type === "TSAbstractMethodDefinition"
    ) {
      if (node.kind === "method" || node.kind === "get")
        throw new Error(`${file}: command registry shape cannot be censused`);
      return;
    }
    // Destructuring patterns bind names; only object literals register commands.
    if (
      node.type !== "Property" ||
      ancestors.at(-1)?.type !== "ObjectExpression"
    )
      return;
    if (node.shorthand || node.method || node.kind === "get")
      throw new Error(`${file}: command registry shape cannot be censused`);
    if (node.kind !== "init") return;
    const object = unwrapExpression(node.value);
    if (object.type !== "ObjectExpression")
      throw new Error(
        `${file}: command registry must be an object literal for the census`,
      );
    for (const command of object.properties) {
      const name = propertyName(command);
      const initializer = commandInitializer(command);
      if (!name || !initializer)
        throw new Error(
          `${file}: command registration cannot be censused: ${nodeText(text, command)}`,
        );
      registrations.push({ file, name, initializer });
    }
  });
  return registrations;
};

/** A variable's initializer; null for a variable without one. */
const declarationValue = (declaration: Declaration): Node | null => {
  if (declaration.node.type === "VariableDeclarator")
    return declaration.node.init;
  if (declaration.node.type === "CatchClause") return null;
  return declaration.node;
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
    const { program } = parseSourceModule(normalizedFile, source);
    files.set(normalizedFile, {
      program,
      text: source,
      declarations: collectDeclarations(program, normalizedFile),
      imports: importedBindings(program),
    });
  }

  const resolveDeclaration = (
    file: string,
    name: string,
  ): Declaration | undefined => {
    const fileIndex = files.get(file);
    if (!fileIndex) return undefined;
    if (fileIndex.imports.has(name)) {
      const imported = fileIndex.imports.get(name);
      if (!imported) return undefined;
      const importedFile = resolveModuleFile(file, imported.source, files);
      return importedFile
        ? (files.get(importedFile)?.declarations.get(imported.name) ??
            undefined)
        : undefined;
    }
    return fileIndex.declarations.get(name) ?? undefined;
  };

  const isCanonicalWrapper = (file: string, name: string): boolean => {
    const imported = files.get(file)?.imports.get(name);
    if (!imported || !canonicalWrapperNames.has(imported.name)) return false;
    return (
      resolveModuleFile(file, imported.source, files) === canonicalCommandsFile
    );
  };

  const proveDeclaration = (
    declaration: Declaration,
    visited: Set<string>,
  ): boolean => {
    const declarationNameValue = declarationName(declaration.node);
    const key = `${declaration.file}#${declarationNameValue ?? declaration.node.start}`;
    if (visited.has(key)) return false;
    const nextVisited = new Set(visited);
    nextVisited.add(key);

    const value = declarationValue(declaration);
    if (
      declaration.node.type === "VariableDeclarator" ||
      declaration.node.type === "CatchClause"
    ) {
      if (value === null) return false;
      const initializer = unwrapExpression(value);
      if (
        initializer.type === "ArrowFunctionExpression" ||
        initializer.type === "FunctionExpression"
      ) {
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

    const returns = value === null ? [] : returnedExpressions(value);
    return (
      returns.length > 0 &&
      returns.every(
        ({ expression }) =>
          expression !== undefined &&
          proveExpression(declaration.file, expression, nextVisited),
      )
    );
  };

  const proveExpression = (
    file: string,
    input: Node,
    visited: Set<string>,
  ): boolean => {
    const expression = unwrapExpression(input);
    if (
      expression.type === "CallExpression" &&
      unwrapExpression(expression.callee).type === "Identifier"
    ) {
      const callee = unwrapExpression(expression.callee);
      if (callee.type === "Identifier" && isCanonicalWrapper(file, callee.name))
        return true;
      if (callee.type === "Identifier") {
        const declaration = resolveDeclaration(file, callee.name);
        return (
          declaration !== undefined && proveDeclaration(declaration, visited)
        );
      }
      return false;
    }

    if (expression.type === "Identifier") {
      if (isCanonicalWrapper(file, expression.name)) return true;
      const declaration = resolveDeclaration(file, expression.name);
      return (
        declaration !== undefined && proveDeclaration(declaration, visited)
      );
    }

    if (expression.type === "ConditionalExpression") {
      return (
        proveExpression(file, expression.consequent, visited) &&
        proveExpression(file, expression.alternate, visited)
      );
    }
    return false;
  };

  const provesFactory = (file: string, initializer: Node): boolean => {
    if (unwrapExpression(initializer).type === "Identifier") {
      return proveExpression(file, initializer, new Set());
    }
    const returns = returnedExpressions(initializer);
    return (
      returns.length > 0 &&
      returns.every(
        ({ expression }) =>
          expression !== undefined &&
          proveExpression(file, expression, new Set()),
      )
    );
  };

  const registrations = sources.flatMap(({ file }) => {
    const normalizedFile = path.posix.normalize(file);
    const fileIndex = files.get(normalizedFile);
    return fileIndex
      ? commandObjectRegistrations(
          fileIndex.program,
          fileIndex.text,
          normalizedFile,
        )
      : [];
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
  for (const command of baseline)
    available.set(command, (available.get(command) ?? 0) + 1);
  const additions: string[] = [];
  for (const command of current) {
    const count = available.get(command) ?? 0;
    if (count === 0) additions.push(command);
    else available.set(command, count - 1);
  }
  return additions;
};
