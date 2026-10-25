// Syntax-only source analysis for repository scripts and source guards.
//
// Parses with oxc (ESTree shape, no parent links). A file with a syntax error
// stops the caller: a guard that read half a tree would report a clean result
// for code it never saw.
import {
  type ArrayPattern,
  type Comment,
  type EcmaScriptModule,
  type Expression,
  type Node,
  type ObjectPattern,
  type ParserOptions,
  parseSync,
  type Program,
  type TSType,
  visitorKeys,
} from "oxc-parser";

export type { Comment, EcmaScriptModule, Node, Program };

export type SourceModule = {
  readonly fileName: string;
  readonly text: string;
  readonly program: Program;
  readonly module: EcmaScriptModule;
  readonly comments: readonly Comment[];
};

/** Parse a module; the language follows the file extension unless `options.lang` names it. */
export const parseSourceModule = (
  fileName: string,
  text: string,
  options: ParserOptions = {},
): SourceModule => {
  const result = parseSync(fileName, text, options);
  const [error] = result.errors;
  if (error !== undefined) {
    throw new Error(`${fileName}: ${error.message}`);
  }
  return {
    fileName,
    text,
    program: result.program,
    module: result.module,
    comments: result.comments,
  };
};

const isNode = (value: unknown): value is Node =>
  typeof value === "object" &&
  value !== null &&
  typeof Reflect.get(value, "type") === "string";

const childrenOf = (node: Node): Node[] => {
  const children: Node[] = [];
  for (const key of visitorKeys[node.type] ?? []) {
    const child: unknown = Reflect.get(node, key);
    if (Array.isArray(child)) {
      for (const item of child) {
        if (isNode(item)) children.push(item);
      }
    } else if (isNode(child)) {
      children.push(child);
    }
  }
  return children;
};

/**
 * Depth-first, source-order walk. `ancestors` runs from the root to the
 * node's parent. Returning `false` skips the node's children.
 */
export const walkAst = (
  root: Node,
  visit: (node: Node, ancestors: readonly Node[]) => boolean | void,
): void => {
  const ancestors: Node[] = [];
  const step = (node: Node): void => {
    if (visit(node, ancestors) === false) return;
    ancestors.push(node);
    for (const child of childrenOf(node)) step(child);
    ancestors.pop();
  };
  step(root);
};

/** Zero-based line and column (UTF-16 units) of an offset, as editors count. */
export const createLineIndex = (
  text: string,
): ((offset: number) => { line: number; column: number }) => {
  const starts = [0];
  for (let index = 0; index < text.length; index += 1) {
    if (text.charCodeAt(index) === 10) starts.push(index + 1);
  }
  return (offset) => {
    let low = 0;
    let high = starts.length - 1;
    while (low < high) {
      const middle = (low + high + 1) >> 1;
      if ((starts[middle] ?? 0) <= offset) low = middle;
      else high = middle - 1;
    }
    return { line: low, column: offset - (starts[low] ?? 0) };
  };
};

/** The source text a node spans. */
export const nodeText = (text: string, node: { start: number; end: number }) =>
  text.slice(node.start, node.end);

// TypeScript-shaped reading. A guard written against TypeScript's syntax tree
// keeps its meaning when the tree it walks has the same nodes; these helpers
// absorb the places where ESTree adds, splits or repeats one.

/**
 * Parse as TypeScript's `createSourceFile` reads an explicit script kind: the
 * grammar follows `lang`, not the extension, and both a top-level `await` and
 * a top-level `return` parse. oxc accepts the latter only as CommonJS, which
 * rejects the former, so CommonJS is the fallback and the module error is the
 * one reported. As in TypeScript, a `.d.ts` name makes the file a declaration
 * file whatever the script kind. An optional chain is no node of its own in TypeScript, so each
 * `ChainExpression` is replaced in place by the expression it wraps (same
 * span): a guard asking whether an initializer is a call then sees the call.
 */
export const parseLikeTypeScript = (
  fileName: string,
  text: string,
  kind: "ts" | "tsx" | "jsx",
): SourceModule => {
  const lang = kind === "ts" && /\.d\.[cm]?ts$/u.test(fileName) ? "dts" : kind;
  let parsed: SourceModule;
  try {
    parsed = parseSourceModule(fileName, text, { lang });
  } catch (moduleError) {
    const script = parseSync(fileName, text, { lang, sourceType: "commonjs" });
    if (script.errors.length > 0) throw moduleError;
    parsed = {
      fileName,
      text,
      program: script.program,
      module: script.module,
      comments: script.comments,
    };
  }
  walkAst(parsed.program, (node) => {
    for (const key of visitorKeys[node.type] ?? []) {
      const child: unknown = Reflect.get(node, key);
      if (Array.isArray(child)) {
        for (let index = 0; index < child.length; index += 1) {
          const item: unknown = child[index];
          if (isNode(item) && item.type === "ChainExpression")
            child[index] = item.expression;
        }
      } else if (isNode(child) && child.type === "ChainExpression") {
        Reflect.set(node, key, child.expression);
      }
    }
  });
  return parsed;
};

/** ESTree spells `{ a }`, `import { a }` and `export { a }` with two identifiers; TypeScript has one. */
const isRepeatedName = (node: Node, parent: Node | undefined): boolean => {
  if (parent === undefined || node.type !== "Identifier") return false;
  if (parent.type === "Property")
    return parent.shorthand && parent.key === node;
  if (parent.type === "ImportSpecifier")
    return parent.imported === node && parent.local.start === node.start;
  if (parent.type === "ExportSpecifier")
    return parent.exported === node && parent.local.start === node.start;
  return false;
};

export type SyntaxTree = {
  /** Every node once, in source order. */
  readonly nodes: readonly Node[];
  readonly parentOf: (node: Node) => Node | undefined;
};

export const collectSyntaxTree = (root: Node): SyntaxTree => {
  const nodes: Node[] = [];
  const parents = new Map<Node, Node>();
  walkAst(root, (node, ancestors) => {
    const parent = ancestors.at(-1);
    if (isRepeatedName(node, parent)) return false;
    nodes.push(node);
    if (parent !== undefined) parents.set(node, parent);
  });
  return { nodes, parentOf: (node) => parents.get(node) };
};

/**
 * TypeScript's `isFunctionLike`. A class or object method is the
 * `FunctionExpression` under its `MethodDefinition` or `Property`.
 */
const FUNCTION_LIKE = new Set<string>([
  "FunctionDeclaration",
  "FunctionExpression",
  "ArrowFunctionExpression",
  "TSDeclareFunction",
  "TSEmptyBodyFunctionExpression",
  "TSFunctionType",
  "TSConstructorType",
  "TSMethodSignature",
  "TSCallSignatureDeclaration",
  "TSConstructSignatureDeclaration",
  "TSIndexSignature",
]);

export const isFunctionLike = (node: Node): boolean =>
  FUNCTION_LIKE.has(node.type);

/** TypeScript's `isBlock`; a class static block is its own block here. */
export const isBlock = (node: Node): boolean =>
  node.type === "BlockStatement" || node.type === "StaticBlock";

/** The parameter list of a function-like node (empty for anything else). */
export const parametersOf = (node: Node): readonly Node[] => {
  switch (node.type) {
    case "FunctionDeclaration":
    case "FunctionExpression":
    case "ArrowFunctionExpression":
    case "TSDeclareFunction":
    case "TSEmptyBodyFunctionExpression":
    case "TSFunctionType":
    case "TSConstructorType":
    case "TSMethodSignature":
    case "TSCallSignatureDeclaration":
    case "TSConstructSignatureDeclaration":
      return node.params;
    case "TSIndexSignature":
      return node.parameters;
    default:
      return [];
  }
};

const annotationOf = (node: Node): TSType | undefined => {
  switch (node.type) {
    case "Identifier":
    case "ObjectPattern":
    case "ArrayPattern":
    case "RestElement":
      return node.typeAnnotation?.typeAnnotation;
    default:
      return undefined;
  }
};

/**
 * A declared name as TypeScript splits it. `name` is the bound identifier or
 * pattern; read an identifier's `.name`, since its ESTree span also covers
 * the type annotation.
 */
export type DeclaredName = {
  readonly name: Node;
  readonly type: TSType | undefined;
  readonly initializer: Expression | undefined;
};

/** One entry of `parametersOf`. ESTree hangs a parameter's type on its binding name. */
export const parameterParts = (parameter: Node): DeclaredName => {
  const target =
    parameter.type === "TSParameterProperty" ? parameter.parameter : parameter;
  if (target.type === "AssignmentPattern")
    return {
      name: target.left,
      type: annotationOf(target.left),
      initializer: target.right,
    };
  return {
    name: target.type === "RestElement" ? target.argument : target,
    type: annotationOf(target),
    initializer: undefined,
  };
};

/** TypeScript's `BindingElement`; `node` is the pattern property, array slot or rest element. */
export type BindingElement = DeclaredName & {
  readonly node: Node;
  /** The written key of `{ key: name }`; absent for shorthand, array and rest entries. */
  readonly propertyName: Node | undefined;
};

const withDefault = (
  target: Node,
): { name: Node; initializer: Expression | undefined } =>
  target.type === "AssignmentPattern"
    ? { name: target.left, initializer: target.right }
    : { name: target, initializer: undefined };

/** The elements of an object or array binding pattern; array holes are skipped. */
export const bindingElements = (pattern: Node): BindingElement[] => {
  let entries: ReadonlyArray<
    ObjectPattern["properties"][number] | ArrayPattern["elements"][number]
  > = [];
  if (pattern.type === "ObjectPattern") entries = pattern.properties;
  if (pattern.type === "ArrayPattern") entries = pattern.elements;
  const elements: BindingElement[] = [];
  for (const entry of entries) {
    if (entry === null) continue;
    if (entry.type === "RestElement")
      elements.push({
        node: entry,
        name: entry.argument,
        type: undefined,
        initializer: undefined,
        propertyName: undefined,
      });
    else
      elements.push({
        node: entry,
        ...withDefault(entry.type === "Property" ? entry.value : entry),
        type: undefined,
        propertyName:
          entry.type === "Property" && !entry.shorthand ? entry.key : undefined,
      });
  }
  return elements;
};

/** TypeScript's `propertyName.getText()`: a computed key keeps its brackets. */
export const propertyNameText = (
  text: string,
  element: BindingElement,
): string | undefined => {
  const key = element.propertyName;
  if (key === undefined) return undefined;
  if (element.node.type === "Property" && element.node.computed)
    return text.slice(element.node.start, text.indexOf("]", key.end) + 1);
  return nodeText(text, key);
};

/** The value of a string literal or of a template literal without substitutions. */
export const staticStringValue = (node: Node): string | undefined => {
  if (node.type === "Literal") {
    const { value } = node;
    return typeof value === "string" ? value : undefined;
  }
  if (node.type === "TemplateLiteral" && node.expressions.length === 0)
    return node.quasis.at(0)?.value.cooked ?? undefined;
  return undefined;
};

export type ModuleReference = {
  readonly specifier: string;
  readonly typeOnly: boolean;
};

/**
 * Every module a file names with a static specifier, in source order: import
 * and re-export declarations, `import x = require()`, `import("…")` types,
 * and dynamic `import()` and `require()` calls.
 */
export const moduleReferences = (program: Program): ModuleReference[] => {
  const references: ModuleReference[] = [];
  const add = (node: Node | undefined, typeOnly: boolean) => {
    const specifier = node === undefined ? undefined : staticStringValue(node);
    if (specifier !== undefined) references.push({ specifier, typeOnly });
  };
  walkAst(program, (node) => {
    if (node.type === "ImportDeclaration") {
      add(
        node.source,
        node.importKind === "type" ||
          (node.specifiers.length > 0 &&
            node.specifiers.every(
              (specifier) =>
                specifier.type === "ImportSpecifier" &&
                specifier.importKind === "type",
            )),
      );
    } else if (node.type === "ExportNamedDeclaration") {
      if (node.source !== null)
        add(
          node.source,
          node.exportKind === "type" ||
            (node.specifiers.length > 0 &&
              node.specifiers.every(
                (specifier) => specifier.exportKind === "type",
              )),
        );
    } else if (node.type === "ExportAllDeclaration") {
      add(node.source, node.exportKind === "type");
    } else if (node.type === "TSImportEqualsDeclaration") {
      if (node.moduleReference.type === "TSExternalModuleReference")
        add(node.moduleReference.expression, node.importKind === "type");
    } else if (node.type === "TSImportType") {
      add(node.source, true);
    } else if (node.type === "ImportExpression") {
      add(node.source, false);
    } else if (
      node.type === "CallExpression" &&
      node.callee.type === "Identifier" &&
      node.callee.name === "require"
    ) {
      add(node.arguments.at(0), false);
    }
  });
  return references;
};

/**
 * Every module a file names, in source order, as TypeScript's `preProcessFile`
 * lists them: imports and re-exports (type-only included), `import x =
 * require()`, `import()` with a literal specifier and `import("x")` types.
 * Ambient `declare module "x"` names follow when the file is a module.
 */
export const importedModuleNames = ({
  program,
  module,
}: SourceModule): string[] => {
  const specifiers: string[] = [];
  const ambient: string[] = [];
  walkAst(program, (node) => {
    if (
      (node.type === "ImportDeclaration" ||
        node.type === "ExportNamedDeclaration" ||
        node.type === "ExportAllDeclaration") &&
      node.source
    ) {
      specifiers.push(node.source.value);
    }
    if (
      node.type === "TSImportEqualsDeclaration" &&
      node.moduleReference.type === "TSExternalModuleReference"
    ) {
      specifiers.push(node.moduleReference.expression.value);
    }
    if (node.type === "TSImportType") specifiers.push(node.source.value);
    if (node.type === "ImportExpression") {
      const argument = node.source;
      if (argument.type === "Literal" && typeof argument.value === "string") {
        specifiers.push(argument.value);
      }
      if (
        argument.type === "TemplateLiteral" &&
        argument.expressions.length === 0
      ) {
        specifiers.push(argument.quasis[0]?.value.cooked ?? "");
      }
    }
    if (
      node.type === "TSModuleDeclaration" &&
      node.declare &&
      node.id.type === "Literal"
    ) {
      ambient.push(node.id.value);
    }
  });
  return module.hasModuleSyntax ? [...specifiers, ...ambient] : specifiers;
};

/** The member a `a.name` or `a["name"]` access reads; other accesses have none. */
export const staticMemberName = (node: Node): string | undefined => {
  if (node.type !== "MemberExpression") return undefined;
  const { property } = node;
  if (!node.computed) {
    if (property.type === "PrivateIdentifier") return `#${property.name}`;
    return property.type === "Identifier" ? property.name : undefined;
  }
  if (property.type !== "Literal") return undefined;
  const { value } = property;
  return typeof value === "string" ? value : undefined;
};

const PATTERN_PARTS = new Set([
  "ArrayPattern",
  "AssignmentPattern",
  "ObjectPattern",
  "Property",
  "RestElement",
]);

/**
 * Whether the destructuring pattern that `ancestors` ends in declares bindings
 * (a declaration, parameter or catch clause). ESTree also writes destructuring
 * assignments such as `({ a } = b)` as patterns; those assign, not declare.
 */
export const declaresBindings = (ancestors: readonly Node[]): boolean => {
  const owner = ancestors.findLast((node) => !PATTERN_PARTS.has(node.type));
  return !(
    owner?.type === "AssignmentExpression" ||
    owner?.type === "ForInStatement" ||
    owner?.type === "ForOfStatement"
  );
};
