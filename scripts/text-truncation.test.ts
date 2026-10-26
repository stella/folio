import { describe, expect, setDefaultTimeout, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";
import type { TSType } from "oxc-parser";

import { scanTestTimeout } from "../test/scan-timeout";
import {
  bindingElements,
  collectSyntaxTree,
  createLineIndex,
  isBlock,
  isFunctionLike,
  type Node,
  nodeText,
  parameterParts,
  parametersOf,
  parseLikeTypeScript,
  propertyNameText,
  walkAst,
} from "./lib/oxc-ast";

setDefaultTimeout(scanTestTimeout(30_000));

const PACKAGES = path.resolve(import.meta.dir, "../packages");
const OWNER = "utf8-boundaries/src/sliceText.ts";
const EXTENSIONS = [
  "ts",
  "tsx",
  "js",
  "jsx",
  "mts",
  "cts",
  "mjs",
  "cjs",
  "vue",
];
const SOURCE_GLOB = `*/src/**/*.{${EXTENSIONS.join(",")}}`;
const SOURCE_FILES = [...new Bun.Glob(SOURCE_GLOB).scanSync({ cwd: PACKAGES })];
const CAP_NAME =
  /max|limit|radius|preview|snippet|context|truncate|unchanged.*edge/iu;
const TEXT_NAME =
  /^(?:text|message|name|value|json|label|description|title|raw|query|author|snippet|context)$/iu;

type Violation = { line: number; reason: string };

const isStringLiteral = (
  node: Node | undefined,
): node is Node & { type: "Literal"; value: string } =>
  node?.type === "Literal" && typeof node.value === "string";

const isNumericLiteral = (
  node: Node | undefined,
): node is Node & { type: "Literal"; value: number } =>
  node?.type === "Literal" && typeof node.value === "number";

const isRegExpLiteral = (node: Node | undefined): boolean =>
  node?.type === "Literal" && "regex" in node;

/** TypeScript's prefix unary operators; `typeof`, `void` and `delete` are nodes of their own there. */
const isPrefixUnary = (
  node: Node,
): node is Node & { type: "UnaryExpression" } =>
  node.type === "UnaryExpression" &&
  (node.operator === "-" ||
    node.operator === "+" ||
    node.operator === "!" ||
    node.operator === "~");

/** A statement the way TypeScript lists it: an exported declaration keeps its `export`. */
const declarationOf = (statement: Node): Node =>
  (statement.type === "ExportNamedDeclaration" ||
    statement.type === "ExportDefaultDeclaration") &&
  statement.declaration
    ? statement.declaration
    : statement;

const SPACE_INSPECTOR =
  "core/src/layout-engine/measure/paragraphMeasureShared.ts";
const verifiedSpaceInspector = (text: string): boolean => {
  const { program } = parseLikeTypeScript(SPACE_INSPECTOR, text, "ts");
  const declaration = program.body
    .map(declarationOf)
    .find(
      (node) =>
        (node.type === "FunctionDeclaration" ||
          node.type === "TSDeclareFunction") &&
        node.id?.name === "isSpaceOrTab",
    );
  if (
    !declaration ||
    (declaration.type !== "FunctionDeclaration" &&
      declaration.type !== "TSDeclareFunction") ||
    declaration.params.length !== 1 ||
    declaration.body?.body.length !== 1
  )
    return false;
  const parameter = parameterParts(declaration.params[0]!).name;
  const statement = declaration.body.body.at(0);
  if (
    parameter.type !== "Identifier" ||
    !statement ||
    statement.type !== "ReturnStatement" ||
    !statement.argument ||
    statement.argument.type !== "LogicalExpression" ||
    statement.argument.operator !== "||"
  )
    return false;
  const compares = (node: Node, expected: string): boolean =>
    node.type === "BinaryExpression" &&
    node.operator === "===" &&
    node.left.type === "Identifier" &&
    node.left.name === parameter.name &&
    isStringLiteral(node.right) &&
    node.right.value === expected;
  return (
    compares(statement.argument.left, " ") &&
    compares(statement.argument.right, "\t")
  );
};

const member = (node: Node): string | undefined => {
  if (node.type !== "MemberExpression") return undefined;
  if (!node.computed)
    return node.property.type === "PrivateIdentifier"
      ? `#${node.property.name}`
      : node.property.name;
  if (isStringLiteral(node.property)) return node.property.value;
  return undefined;
};

const receiver = (node: Node): Node | undefined =>
  node.type === "MemberExpression" ? node.object : undefined;

/** `a.b` with an identifier name: TypeScript's property access. */
const accessedName = (node: Node): string | undefined =>
  node.type === "MemberExpression" && !node.computed ? member(node) : undefined;

/** A TypeScript parameter or variable declaration. */
type Declared = {
  readonly kind: "parameter" | "variable";
  readonly name: Node;
  readonly type: TSType | undefined;
};

const findTruncations = (input: string, file = "fixture.ts"): Violation[] => {
  const text = file.endsWith(".vue")
    ? [...input.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/gu)]
        .map((match) => match[1] ?? "")
        .join("\n")
    : input;
  const { program } = parseLikeTypeScript(
    file,
    text,
    file.endsWith(".tsx") || file.endsWith(".jsx") ? "tsx" : "ts",
  );
  const { nodes, parentOf } = collectSyntaxTree(program);
  const lineOf = createLineIndex(text);
  // Declarations are keyed by the node that stands where TypeScript's does:
  // the declarator, the parameter, or the catch clause's variable.
  const declared = new Map<Node, Declared>();
  for (const node of nodes) {
    if (node.type === "VariableDeclarator")
      declared.set(node, { kind: "variable", ...parameterParts(node.id) });
    if (node.type === "CatchClause" && node.param)
      declared.set(node.param, {
        kind: "variable",
        ...parameterParts(node.param),
      });
    if (isFunctionLike(node))
      for (const parameter of parametersOf(node))
        declared.set(parameter, {
          kind: "parameter",
          ...parameterParts(parameter),
        });
  }
  // Resolve identifiers to declarations rather than sharing facts by spelling.
  const bindings = new Map<Node, Map<string, Node>>();
  const isScope = (node: Node): boolean =>
    node.type === "Program" ||
    isBlock(node) ||
    isFunctionLike(node) ||
    node.type === "ForStatement" ||
    node.type === "ForOfStatement" ||
    node.type === "ForInStatement" ||
    node.type === "CatchClause" ||
    // TypeScript's case block holds the clauses, not the switch's subject.
    node.type === "SwitchStatement";
  /** TypeScript's `scopeOf(node.parent)`. */
  const scopeAbove = (node: Node): Node => {
    let child = node;
    let current = parentOf(node);
    while (current) {
      if (
        isScope(current) &&
        !(current.type === "SwitchStatement" && current.discriminant === child)
      )
        return current;
      child = current;
      current = parentOf(current);
    }
    return program;
  };
  const bind = (name: Node, declaration: Node, scope: Node): void => {
    if (name.type !== "Identifier") {
      for (const element of bindingElements(name))
        bind(element.name, element.node, scope);
      return;
    }
    let names = bindings.get(scope);
    if (!names) {
      names = new Map();
      bindings.set(scope, names);
    }
    names.set(name.name, declaration);
  };
  for (const node of nodes) {
    const declaration = declared.get(node);
    if (declaration?.kind === "parameter")
      bind(declaration.name, node, scopeAbove(node));
    if (declaration?.kind === "variable") {
      let scope = scopeAbove(node);
      const list = parentOf(node);
      if (list?.type === "VariableDeclaration" && list.kind === "var") {
        for (
          let current: Node | undefined = list;
          current;
          current = parentOf(current)
        ) {
          if (isFunctionLike(current) || current.type === "Program") {
            scope = current;
            break;
          }
        }
      }
      bind(declaration.name, node, scope);
    }
    if (
      (node.type === "FunctionDeclaration" ||
        node.type === "TSDeclareFunction") &&
      node.id
    )
      bind(node.id, node, scopeAbove(node));
    if (
      node.type === "ImportSpecifier" ||
      node.type === "ImportDefaultSpecifier" ||
      node.type === "ImportNamespaceSpecifier"
    )
      bind(node.local, node, scopeAbove(node));
  }
  const bindingOf = (identifier: Node & { name: string }): Node => {
    for (
      let scope: Node | undefined = scopeAbove(identifier);
      scope;
      scope = scope === program ? undefined : scopeAbove(scope)
    ) {
      const declaration = bindings.get(scope)?.get(identifier.name);
      if (declaration) return declaration;
    }
    return identifier;
  };
  const strings = new Set<Node>();
  const arrays = new Set<Node>();
  const stringArrays = new Set<Node>();
  const stringFunctions = new Set<Node>();
  const methods = new Map<Node, { receiver: Node; method: string }>();
  const caps = new Set<Node>();
  const numbers = new Set<Node>();
  const asciiValidators = new Set<Node>();
  const regexes = new Set<Node>();
  const inspectors = new Set<Node>();
  const mutated = new Set<Node>();
  for (const node of nodes) {
    if (node.type === "UpdateExpression" && node.argument.type === "Identifier")
      mutated.add(bindingOf(node.argument));
    if (node.type === "AssignmentExpression" && node.left.type === "Identifier")
      mutated.add(bindingOf(node.left));
  }
  const numericCall = (node: Node): boolean => {
    if (node.type !== "CallExpression") return false;
    if (node.callee.type === "Identifier")
      return (
        /^(?:parseInt|parseFloat|Number)$/u.test(node.callee.name) &&
        bindingOf(node.callee) === node.callee
      );
    const owner = receiver(node.callee);
    return (
      owner !== undefined &&
      owner.type === "Identifier" &&
      owner.name === "Number" &&
      bindingOf(owner) === owner &&
      /^(?:parseInt|parseFloat)$/u.test(member(node.callee) ?? "")
    );
  };
  const typeText = (type: TSType | undefined): string | undefined =>
    type === undefined ? undefined : nodeText(text, type);
  for (const node of nodes) {
    const declaration = declared.get(node);
    if (declaration && declaration.name.type === "Identifier") {
      const { type } = declaration;
      if (type?.type === "TSStringKeyword") strings.add(node);
      if (type?.type === "TSNumberKeyword") numbers.add(node);
      if (typeText(type) === "RegExp") regexes.add(node);
      if (/^RegExp(?:Match|Exec)Array$/u.test(typeText(type) ?? ""))
        arrays.add(node);
      if (
        type &&
        (type.type === "TSArrayType" ||
          /^(?:readonly\s+)?(?:Uint\d+Array|Int\d+Array|ArrayBuffer|Buffer|Array<|ReadonlyArray<)/u.test(
            nodeText(text, type),
          ))
      )
        arrays.add(node);
      if (
        type &&
        type.type === "TSArrayType" &&
        type.elementType.type === "TSStringKeyword"
      )
        stringArrays.add(node);
    }
    if (
      (node.type === "FunctionDeclaration" ||
        node.type === "TSDeclareFunction") &&
      node.id &&
      node.returnType?.typeAnnotation.type === "TSStringKeyword"
    )
      stringFunctions.add(node);
    if (node.type !== "VariableDeclarator") {
      if (
        node.type === "ImportSpecifier" &&
        (node.imported.type === "Literal"
          ? node.imported.value
          : node.imported.name) === "isSpaceOrTab"
      ) {
        const importDeclaration = parentOf(node);
        if (
          importDeclaration?.type === "ImportDeclaration" &&
          path.resolve(
            PACKAGES,
            path.dirname(file),
            `${importDeclaration.source.value}.ts`,
          ) === path.join(PACKAGES, SPACE_INSPECTOR) &&
          verifiedSpaceInspector(
            readFileSync(path.join(PACKAGES, SPACE_INSPECTOR), "utf8"),
          )
        )
          inspectors.add(node);
      }
      continue;
    }
    const initializer = node.init;
    if (
      node.id.type === "Identifier" &&
      initializer?.type === "ArrowFunctionExpression" &&
      initializer.returnType?.typeAnnotation.type === "TSStringKeyword"
    )
      stringFunctions.add(node);
    if (
      node.id.type === "Identifier" &&
      initializer &&
      isRegExpLiteral(initializer) &&
      /^\/\^\[[\da-zA-Z-]+\]\{\d+\}\$\/[gu]*$/u.test(
        nodeText(text, initializer),
      )
    )
      asciiValidators.add(node);
    if (
      initializer &&
      (isRegExpLiteral(initializer) ||
        (initializer.type === "NewExpression" &&
          initializer.callee.type === "Identifier" &&
          initializer.callee.name === "RegExp" &&
          bindingOf(initializer.callee) === initializer.callee))
    )
      regexes.add(node);
    if (
      initializer?.type === "ArrowFunctionExpression" &&
      initializer.body.type !== "BlockStatement" &&
      numericCall(initializer.body)
    )
      inspectors.add(node);
  }
  const isArray = (node: Node): boolean => {
    if (node.type === "ArrayExpression") return true;
    if (node.type === "Identifier") return arrays.has(bindingOf(node));
    if (node.type === "NewExpression")
      return /^(?:Uint\d+Array|Int\d+Array|ArrayBuffer|Buffer|Array)$/u.test(
        nodeText(text, node.callee),
      );
    if (node.type === "CallExpression") {
      if (
        member(node.callee) === "split" ||
        member(node.callee) === "map" ||
        member(node.callee) === "filter"
      )
        return true;
      const owner = receiver(node.callee);
      if (
        member(node.callee) === "from" &&
        owner !== undefined &&
        nodeText(text, owner) === "Array"
      )
        return true;
      if (
        member(node.callee) === "exec" &&
        owner &&
        (isRegExpLiteral(owner) ||
          (owner.type === "Identifier" && regexes.has(bindingOf(owner))))
      )
        return true;
      if (member(node.callee) === "match" && owner && isString(owner))
        return true;
    }
    return false;
  };
  const isStringArray = (node: Node): boolean => {
    if (node.type === "Identifier") return stringArrays.has(bindingOf(node));
    if (node.type === "ArrayExpression")
      return node.elements.every(
        (element) => element !== null && isString(element),
      );
    if (node.type !== "CallExpression") return false;
    const owner = receiver(node.callee);
    if (!owner) return false;
    if (member(node.callee) === "split") return isString(owner);
    return member(node.callee) === "filter" && isStringArray(owner);
  };
  const isString = (node: Node): boolean => {
    if (isArray(node)) return false;
    if (isStringLiteral(node) || node.type === "TemplateLiteral") return true;
    if (node.type === "Identifier")
      return (
        !numbers.has(bindingOf(node)) &&
        (strings.has(bindingOf(node)) || TEXT_NAME.test(node.name))
      );
    if (node.type === "MemberExpression")
      return /^(?:text|message|name|title|description|label)$/u.test(
        member(node) ?? "",
      );
    if (node.type === "CallExpression") {
      if (node.callee.type === "Identifier")
        return (
          node.callee.name === "String" ||
          stringFunctions.has(bindingOf(node.callee))
        );
      const method = member(node.callee);
      if (method === "join") return true;
      if (method === "toString") return true;
      const owner = receiver(node.callee);
      return (
        owner !== undefined &&
        /^(?:slice|substring|substr|trim|trimStart|trimEnd|replace|replaceAll|toUpperCase|toLowerCase|padStart|padEnd|toString)$/u.test(
          method ?? "",
        ) &&
        isString(owner)
      );
    }
    if (node.type === "BinaryExpression" && node.operator === "+")
      return isString(node.left) || isString(node.right);
    return false;
  };
  let changed = true;
  const coordinateWalk = (declaration: Node): boolean => {
    const list = parentOf(declaration);
    if (
      list?.type === "VariableDeclaration" &&
      parentOf(list)?.type === "ForStatement"
    )
      return true;
    return nodes.some((node) => {
      if (node.type !== "WhileStatement" && node.type !== "ForStatement")
        return false;
      const condition = node.test;
      if (!condition) return false;
      let indexed = false;
      walkAst(condition, (part) => {
        if (
          part.type === "MemberExpression" &&
          part.computed &&
          part.property.type === "Identifier" &&
          bindingOf(part.property) === declaration &&
          isString(part.object)
        )
          indexed = true;
      });
      return indexed;
    });
  };
  while (changed) {
    changed = false;
    for (const node of nodes) {
      const list = parentOf(node);
      const loop = list === undefined ? undefined : parentOf(list);
      if (
        node.type === "VariableDeclarator" &&
        node.id.type === "Identifier" &&
        list?.type === "VariableDeclaration" &&
        loop?.type === "ForOfStatement"
      ) {
        const iterable = loop.right;
        const owner =
          iterable.type === "CallExpression"
            ? receiver(iterable.callee)
            : undefined;
        if (
          iterable.type === "CallExpression" &&
          member(iterable.callee) === "matchAll" &&
          owner &&
          isString(owner) &&
          !arrays.has(node)
        ) {
          arrays.add(node);
          changed = true;
        }
      }
      if (
        node.type === "CallExpression" &&
        /^(?:map|filter|forEach)$/u.test(member(node.callee) ?? "")
      ) {
        const owner = receiver(node.callee);
        const callback = node.arguments.at(0);
        if (
          owner &&
          isStringArray(owner) &&
          callback &&
          (callback.type === "ArrowFunctionExpression" ||
            callback.type === "FunctionExpression")
        ) {
          const first = callback.params.at(0);
          const parameter =
            first === undefined ? undefined : parameterParts(first).name;
          if (
            parameter &&
            parameter.type === "Identifier" &&
            !strings.has(bindingOf(parameter))
          ) {
            strings.add(bindingOf(parameter));
            changed = true;
          }
        }
      }
      if (node.type !== "VariableDeclarator" || !node.init) continue;
      const initializer = node.init;
      if (node.id.type === "ObjectPattern" && isString(initializer)) {
        for (const element of bindingElements(node.id)) {
          if (element.name.type !== "Identifier") continue;
          const method =
            propertyNameText(text, element)?.replace(/["']/gu, "") ??
            element.name.name;
          if (
            /^(?:slice|substring|substr)$/u.test(method) &&
            !methods.has(bindingOf(element.name))
          ) {
            methods.set(bindingOf(element.name), {
              receiver: initializer,
              method,
            });
            changed = true;
          }
        }
      }
      if (node.id.type !== "Identifier") continue;
      const name = node.id;
      if (
        initializer.type === "Identifier" &&
        regexes.has(bindingOf(initializer)) &&
        !regexes.has(node)
      ) {
        regexes.add(node);
        changed = true;
      }
      if (
        initializer.type === "Identifier" &&
        inspectors.has(bindingOf(initializer)) &&
        !inspectors.has(node)
      ) {
        inspectors.add(node);
        changed = true;
      }
      if (
        (isNumericLiteral(initializer) ||
          (initializer.type === "Identifier" &&
            numbers.has(bindingOf(initializer)))) &&
        !numbers.has(node)
      ) {
        numbers.add(node);
        changed = true;
      }
      if (isArray(initializer) && !arrays.has(bindingOf(name))) {
        arrays.add(node);
        changed = true;
      }
      if (isStringArray(initializer) && !stringArrays.has(bindingOf(name))) {
        stringArrays.add(node);
        changed = true;
      }
      if (isString(initializer) && !strings.has(bindingOf(name))) {
        strings.add(node);
        changed = true;
      }
      if (
        list?.type === "VariableDeclaration" &&
        (!mutated.has(node) ||
          !/(?:offset|index|idx|start|end|cursor|position|^i$|^j$)/iu.test(
            name.name,
          ) ||
          !coordinateWalk(node)) &&
        ((isNumericLiteral(initializer) && initializer.value >= 1) ||
          (initializer.type === "Identifier" &&
            caps.has(bindingOf(initializer))))
      ) {
        if (!caps.has(bindingOf(name))) {
          caps.add(bindingOf(name));
          changed = true;
        }
      }
      const owner = receiver(initializer);
      const method = member(initializer);
      if (
        owner &&
        isString(owner) &&
        /^(?:slice|substring|substr)$/u.test(method ?? "") &&
        !methods.has(bindingOf(name))
      ) {
        methods.set(bindingOf(name), {
          receiver: owner,
          method: method ?? "",
        });
        changed = true;
      }
      if (
        initializer.type === "Identifier" &&
        methods.has(bindingOf(initializer)) &&
        !methods.has(bindingOf(name))
      ) {
        methods.set(bindingOf(name), methods.get(bindingOf(initializer))!);
        changed = true;
      }
      if (
        initializer.type === "CallExpression" &&
        member(initializer.callee) === "bind"
      ) {
        const reference = receiver(initializer.callee);
        const boundOwner = reference ? receiver(reference) : undefined;
        const boundMethod = reference ? member(reference) : undefined;
        if (
          boundOwner &&
          isString(boundOwner) &&
          /^(?:slice|substring|substr)$/u.test(boundMethod ?? "") &&
          !methods.has(bindingOf(name))
        ) {
          methods.set(bindingOf(name), {
            receiver: boundOwner,
            method: boundMethod ?? "",
          });
          changed = true;
        }
      }
    }
  }
  const sameBinding = (
    left: Node & { name: string },
    right: Node & { name: string },
  ): boolean => {
    const first = bindingOf(left);
    const second = bindingOf(right);
    return (
      first === second ||
      (first === left && second === right && left.name === right.name)
    );
  };
  const coordinate = (node: Node): boolean => {
    const accessed = accessedName(node);
    if (accessed !== undefined)
      return /^(?:index|offset|position|start|end)$/iu.test(accessed);
    if (
      node.type !== "Identifier" ||
      CAP_NAME.test(node.name) ||
      caps.has(bindingOf(node))
    )
      return false;
    if (
      /(?:offset|index|idx|start|end|cursor|position|break|^at$)/iu.test(
        node.name,
      )
    )
      return true;
    const declaration = bindingOf(node);
    return (
      declaration.type === "VariableDeclarator" &&
      declaration.init?.type === "CallExpression" &&
      /^(?:indexOf|lastIndexOf|search)$/u.test(
        member(declaration.init.callee) ?? "",
      )
    );
  };
  // Small offsets beside parser/selection coordinates retain UTF-16 units.
  const capOperands = (left: Node, right: Node): boolean => {
    if (
      (coordinate(left) && isNumericLiteral(right) && right.value <= 4) ||
      (coordinate(right) && isNumericLiteral(left) && left.value <= 4)
    )
      return false;
    return cap(left) || cap(right);
  };
  const cap = (node: Node | undefined): boolean => {
    if (!node) return false;
    if (isNumericLiteral(node)) return node.value >= 1;
    if (node.type === "Identifier")
      return caps.has(bindingOf(node)) || CAP_NAME.test(node.name);
    if (isPrefixUnary(node) && isNumericLiteral(node.argument))
      return node.argument.value >= 1;
    if (node.type === "CallExpression")
      return node.arguments.some((argument) => cap(argument));
    // TypeScript's binary expression also covers logical, assignment and comma
    // operators; a comma list nests to the left.
    if (
      node.type === "BinaryExpression" ||
      node.type === "LogicalExpression" ||
      node.type === "AssignmentExpression"
    )
      return capOperands(node.left, node.right);
    if (node.type === "SequenceExpression") {
      const [first, second, ...rest] = node.expressions;
      return (
        (first !== undefined &&
          second !== undefined &&
          capOperands(first, second)) ||
        rest.some((expression) => cap(expression))
      );
    }
    const accessed = accessedName(node);
    if (accessed !== undefined) return CAP_NAME.test(accessed);
    return false;
  };
  const isInspection = (node: Node): boolean => {
    // Numeric and regexp predicates inspect units; they do not emit clipped text.
    let parent = parentOf(node);
    while (
      parent?.type === "TSNonNullExpression" ||
      parent?.type === "ParenthesizedExpression"
    )
      parent = parentOf(parent);
    if (parent?.type !== "CallExpression") return false;
    if (numericCall(parent)) return true;
    if (
      parent.callee.type === "Identifier" &&
      inspectors.has(bindingOf(parent.callee))
    )
      return true;
    const owner = receiver(parent.callee);
    return (
      member(parent.callee) === "test" &&
      owner !== undefined &&
      (isRegExpLiteral(owner) ||
        (owner.type === "Identifier" && regexes.has(bindingOf(owner))))
    );
  };
  const asciiTailCondition = (
    condition: Node,
    owner: Node & { name: string },
    amount: number,
    seen = new Set<Node>(),
  ): boolean => {
    if (seen.has(condition)) return false;
    seen.add(condition);
    if (condition.type === "ParenthesizedExpression")
      return asciiTailCondition(condition.expression, owner, amount, seen);
    if (condition.type === "Identifier") {
      const declaration = bindingOf(condition);
      return (
        declaration.type === "VariableDeclarator" &&
        declaration.init !== null &&
        asciiTailCondition(declaration.init, owner, amount, seen)
      );
    }
    if (condition.type === "LogicalExpression") {
      if (condition.operator === "&&")
        return (
          asciiTailCondition(condition.left, owner, amount, seen) ||
          asciiTailCondition(condition.right, owner, amount, seen)
        );
      if (condition.operator === "||")
        return (
          asciiTailCondition(condition.left, owner, amount, new Set(seen)) &&
          asciiTailCondition(condition.right, owner, amount, new Set(seen))
        );
    }
    if (
      condition.type !== "CallExpression" ||
      member(condition.callee) !== "endsWith"
    )
      return false;
    const tested = receiver(condition.callee);
    const suffix = condition.arguments.at(0);
    return (
      tested !== undefined &&
      tested.type === "Identifier" &&
      sameBinding(tested, owner) &&
      isStringLiteral(suffix) &&
      suffix.value.length === amount &&
      suffix.value.length > 0 &&
      Array.from(suffix.value).every(
        (character) => character.charCodeAt(0) <= 0x7f,
      )
    );
  };
  const within = (node: Node, container: Node): boolean =>
    node.start >= container.start && node.end <= container.end;
  const protectedAsciiTail = (
    node: Node & { type: "CallExpression" },
    owner: Node,
  ): boolean => {
    if (owner.type !== "Identifier") return false;
    const end =
      node.arguments.at(1) ??
      (node.arguments.length === 1 ? node.arguments.at(0) : undefined);
    let amount: number | undefined;
    if (
      end &&
      end.type === "UnaryExpression" &&
      end.operator === "-" &&
      isNumericLiteral(end.argument)
    )
      amount = end.argument.value;
    if (
      end &&
      end.type === "BinaryExpression" &&
      end.operator === "-" &&
      accessedName(end.left) === "length" &&
      end.left.type === "MemberExpression" &&
      end.left.object.type === "Identifier" &&
      sameBinding(end.left.object, owner) &&
      isNumericLiteral(end.right)
    )
      amount = end.right.value;
    if (amount === undefined) return false;
    for (let parent = parentOf(node); parent; parent = parentOf(parent)) {
      if (isFunctionLike(parent)) break;
      if (
        parent.type === "IfStatement" &&
        within(node, parent.consequent) &&
        asciiTailCondition(parent.test, owner, amount)
      )
        return true;
      if (
        parent.type === "ConditionalExpression" &&
        within(node, parent.consequent) &&
        asciiTailCondition(parent.test, owner, amount)
      )
        return true;
    }
    return false;
  };
  const asciiNumber = (node: Node): boolean => {
    if (node.type !== "CallExpression") return false;
    const method = member(node.callee);
    const owner = receiver(node.callee);
    if (!owner) return false;
    if (/^(?:toUpperCase|toLowerCase|padStart|padEnd)$/u.test(method ?? "")) {
      const padding = node.arguments.at(1);
      return (
        asciiNumber(owner) &&
        (!padding ||
          (isStringLiteral(padding) &&
            Array.from(padding.value).every(
              (character) => character.charCodeAt(0) <= 0x7f,
            )))
      );
    }
    const radix = node.arguments.at(0);
    if (
      method !== "toString" ||
      !/^(?:8|10|16|36)$/u.test(
        radix === undefined ? "" : nodeText(text, radix),
      )
    )
      return false;
    return (
      isNumericLiteral(owner) ||
      (owner.type === "Identifier" && numbers.has(bindingOf(owner))) ||
      (owner.type === "CallExpression" &&
        owner.callee.type === "MemberExpression" &&
        !owner.callee.computed &&
        owner.callee.object.type === "Identifier" &&
        bindingOf(owner.callee.object) === owner.callee.object &&
        /^(?:Math\.random|Date\.now)$/u.test(nodeText(text, owner.callee)))
    );
  };
  const violations: Violation[] = [];
  const add = (node: Node, reason: string): void => {
    violations.push({
      line: lineOf(node.start).line + 1,
      reason,
    });
  };
  for (const node of nodes) {
    if (
      node.type === "MemberExpression" &&
      node.computed &&
      isNumericLiteral(node.property) &&
      node.property.value === 0 &&
      isString(node.object)
    ) {
      // A following ASCII delimiter proves the first unit is not half a pair.
      let delimiterGuard = false;
      for (let parent = parentOf(node); parent; parent = parentOf(parent)) {
        if (parent.type !== "IfStatement") continue;
        const condition = parent.test;
        if (
          condition.type === "BinaryExpression" &&
          condition.operator === "===" &&
          condition.left.type === "MemberExpression" &&
          condition.left.computed &&
          nodeText(text, condition.left.property) === "1" &&
          isStringLiteral(condition.right) &&
          condition.right.value === ":" &&
          node.object.type === "Identifier" &&
          condition.left.object.type === "Identifier" &&
          sameBinding(node.object, condition.left.object) &&
          within(node, parent.consequent)
        )
          delimiterGuard = true;
        break;
      }
      if (!delimiterGuard && !isInspection(node))
        add(node, "string initials must use grapheme slicing");
    }
    if (node.type !== "CallExpression") continue;
    const alias =
      node.callee.type === "Identifier"
        ? methods.get(bindingOf(node.callee))
        : undefined;
    const owner = alias?.receiver ?? receiver(node.callee);
    const method = alias?.method ?? member(node.callee);
    const first = node.arguments.at(0);
    if (
      owner &&
      isString(owner) &&
      (method === "at" || method === "charAt") &&
      ((isNumericLiteral(first) && first.value === 0) ||
        (method === "charAt" && first === undefined))
    ) {
      if (!asciiNumber(owner) && !isInspection(node))
        add(node, "string initials must use grapheme slicing");
      continue;
    }
    if (
      !owner ||
      !isString(owner) ||
      !/^(?:slice|substring|substr)$/u.test(method ?? "")
    )
      continue;
    if (asciiNumber(owner)) continue;
    if (isInspection(node) || protectedAsciiTail(node, owner)) continue;
    // Parsing a hexadecimal numeric field does not emit the sliced text.
    const parent = parentOf(node);
    const numericParser =
      parent?.type === "CallExpression" &&
      ((parent.callee.type === "Identifier" &&
        parent.callee.name === "parseInt" &&
        bindingOf(parent.callee) === parent.callee) ||
        (accessedName(parent.callee) === "parseInt" &&
          parent.callee.type === "MemberExpression" &&
          parent.callee.object.type === "Identifier" &&
          parent.callee.object.name === "Number" &&
          bindingOf(parent.callee.object) === parent.callee.object));
    const radix =
      parent?.type === "CallExpression" ? parent.arguments.at(1) : undefined;
    if (numericParser && radix !== undefined && nodeText(text, radix) === "16")
      continue;
    if (
      parent?.type === "ConditionalExpression" &&
      parent.consequent === node &&
      parent.test.type === "CallExpression" &&
      member(parent.test.callee) === "test"
    ) {
      const validator = receiver(parent.test.callee);
      const validated = parent.test.arguments.at(0);
      if (
        validator &&
        validator.type === "Identifier" &&
        asciiValidators.has(bindingOf(validator)) &&
        validated &&
        validated.type === "Identifier" &&
        owner.type === "Identifier" &&
        sameBinding(validated, owner)
      )
        continue;
    }
    const only = node.arguments.length === 1 ? node.arguments[0] : undefined;
    if (
      method === "substr" ||
      cap(node.arguments.at(1)) ||
      (only !== undefined && isPrefixUnary(only) && cap(only))
    )
      add(node, "bounded string truncation must use sliceText");
  }
  return violations;
};

// A narrow exception: this function admits only ASCII characters before its cap.
const ASCII_FONT_FUNCTION =
  'constsanitizePostScriptName=(raw:string):string=>{letname="";for(constcharacterofraw){constcode=character.codePointAt(0)??0;if(code>0x20&&code<0x7f&&!POSTSCRIPT_NAME_DELIMITERS.has(character)){name+=character;}}returnname.slice(0,MAX_POSTSCRIPT_NAME_LENGTH);};';
const violationsFor = (file: string, text: string): Violation[] => {
  if (
    file === OWNER ||
    /(?:^|\/)(?:__tests__|__fixtures__|tests?|fixtures?)(?:\/|$)/u.test(file) ||
    /\.(?:test|spec|fixture)\.[cm]?[jt]sx?$/u.test(file)
  )
    return [];
  if (file === "core/src/fonts/sfnt/parse.ts") {
    const { program } = parseLikeTypeScript(file, text, "ts");
    // TypeScript's statement is the whole `export const`, keyword included.
    const ascii = program.body.find((statement) => {
      const declaration = declarationOf(statement);
      return (
        declaration.type === "VariableDeclaration" &&
        declaration.declarations.some(
          ({ id }) =>
            id.type === "Identifier" && id.name === "sanitizePostScriptName",
        )
      );
    });
    if (
      ascii &&
      nodeText(text, ascii).replace(/\s+/gu, "") === ASCII_FONT_FUNCTION
    )
      return findTruncations(
        text.slice(0, ascii.start) +
          " ".repeat(ascii.end - ascii.start) +
          text.slice(ascii.end),
        file,
      );
  }
  return findTruncations(text, file);
};

describe("text truncation boundaries", () => {
  test("every production package and source extension rejects bounded text and initials", async () => {
    const packages = new Set<string>();
    for (const file of SOURCE_FILES) packages.add(file.split("/")[0]!);
    const fixtures = [
      "function preview(text: string) { return text.slice(0, 40); }",
      "const original = '👩🏽‍⚖️'; const alias = original; alias.substring(0, MAX_PREVIEW);",
      "const text = `prefix${value}`; const cut = text.slice; cut(0, maxLength);",
      "const bound = 40; const copiedBound = bound; text.slice(0, copiedBound);",
      "let bound = 40; text.slice(0, bound);",
      "let end = 1; end++; name.slice(0, end);",
      "let offset = 1; while (offset < 3) offset++; name.slice(0, offset);",
      "const end = 40; text.slice(0, end + 1);",
      "const LIMIT = 40; text.slice(0, Math.max(0, LIMIT));",
      "text.slice(0, text.length - 2)",
      "if (text.endsWith('👩')) return text.slice(0, -1);",
      "const emit = { test: (text: string) => text }; emit.test(text.slice(0, 40));",
      "const expand = (part: string) => part; expand(text[0]!);",
      "const Number = { parseInt: (text: string) => text }; Number.parseInt(text.slice(0, 2), 16);",
      "const { slice: cut } = text; const alias = cut; alias(0, LIMIT);",
      "const cut = text.slice.bind(text); cut(0, PREVIEW_LIMIT);",
      "record.message.slice(0, LIMIT);",
      "text.slice(-30)",
      "function initials(name: string) { return name.slice(0, 1); }",
      "function initials(name: string) { return name.slice(-1); }",
      "function initials(name: string) { return name.at(0); }",
      "function initials(name: string) { return name.charAt(0); }",
      "function initials(name: string) { return name.charAt(); }",
      "const bound = 1; name.slice(0, bound);",
      "text.substr(start, length)",
      "function initials(name: string) { return name[0]; }",
      "name.split(/\\s+/).map(word => word[0]).join('').slice(0, 2)",
      "function bytes(text: Uint8Array) { const alias = text; return alias.slice(0, 40); } function preview(text: string) { const alias = text; return alias.slice(0, 40); }",
      "function number(value: number) { return value.toString(16).slice(-8); } function preview(value: string) { return value.toString(16).slice(0, 8); }",
      "const value = 'text'; const copied = value; { const value = new Uint8Array(); const copied = value; copied.slice(0, 40); } copied.slice(0, 40);",
      "function numeric(value: number) { const copied = value; return copied.toString(16).slice(-8); } function preview(value: string) { const copied = value; const cut = copied.slice; const alias = cut; return alias(0, 40); }",
    ];
    for (const packageName of packages)
      for (const extension of EXTENSIONS)
        for (const fixture of fixtures) {
          const file = `${packageName}/src/boundary-fixture.${extension}`;
          const text =
            extension === "vue" ? `<script setup>${fixture}</script>` : fixture;
          expect({
            file,
            fixture,
            rejected: violationsFor(file, text).length > 0,
          }).toEqual({
            file,
            fixture,
            rejected: true,
          });
        }
  });
  test("UTF-16 coordinates, byte buffers, arrays and parsed ASCII numbers retain their units", () => {
    expect(
      findTruncations(
        "function select(text: string, startOffset: number, endOffset: number) { return text.slice(startOffset, endOffset); } text.slice(1); text.charCodeAt(0); const bytes = new Uint8Array(100); bytes.slice(0, LIMIT); bytes[0]; const names: string[] = []; names.slice(0, 20); names[0]; const hex = 'FF'; Number.parseInt(hex.slice(0, 2), 16);",
      ),
    ).toEqual([]);
  });
  test("parser delimiters, mutable coordinate walks and scalar inspections are not previews", () => {
    const fixtures = [
      'function tag(xml: string) { const tagEnd = xml.indexOf(">"); return xml.slice(0, tagEnd + 1); }',
      'function path(text: string) { const lastSlash = text.lastIndexOf("/"); return text.slice(0, lastSlash + 1); }',
      "function whitespace(text: string) { const start = text.length - text.trimStart().length; const end = text.trimEnd().length; return text.slice(start, Math.max(start, end)); }",
      "function partition(text: string) { let start = 0; for (let offset = 1; offset < text.length; offset++) { text.slice(start, offset); start = offset; } }",
      'function period(value: string) { return value.endsWith(".") ? value.slice(0, -1) : value; }',
      'function tag(sourceXml: string) { if (sourceXml.endsWith("/>")) return sourceXml.slice(1, sourceXml.length - 2) + sourceXml.slice(0, -2); return sourceXml; }',
      "const regex = /[\\p{L}]$/u; regex.test(text.slice(0, LIMIT));",
      "const name = /x/u.exec(text); if (name) name[0];",
      "for (const name of text.matchAll(/x/gu)) name[0];",
      "const expand = (part: string): number => Number.parseInt(part.repeat(2), 16); expand(text[0]);",
      "const expand = (part: string): number => Number.parseInt(part.repeat(2), 16); expand(text[0]!);",
      "Number.parseFloat(text.slice(0, -1));",
    ];
    for (const fixture of fixtures)
      expect({ fixture, violations: findTruncations(fixture) }).toEqual({
        fixture,
        violations: [],
      });
  });
  test("the imported space inspector must retain its actual ASCII-only boolean contract", () => {
    const helper = readFileSync(path.join(PACKAGES, SPACE_INSPECTOR), "utf8");
    expect(verifiedSpaceInspector(helper)).toBe(true);
    expect(
      verifiedSpaceInspector(helper.replace('char === " "', 'char === "👩"')),
    ).toBe(false);
    expect(
      findTruncations(
        'import { isSpaceOrTab } from "./paragraphMeasureShared"; isSpaceOrTab(text.at(0)); isSpaceOrTab(text[0]);',
        "core/src/layout-engine/measure/fixture.ts",
      ),
    ).toEqual([]);
    expect(
      findTruncations(
        "const isSpaceOrTab = (text: string) => text; isSpaceOrTab(text[0]);",
      ),
    ).toHaveLength(1);
  });
  test("declaration identities isolate shadowed string, array, numeric and copied method bindings", () => {
    const source =
      "function bytes(text: Uint8Array) { const alias = text; return alias.slice(0, 40); } function preview(text: string) { const alias = text; return alias.slice(0, 40); }";
    expect(findTruncations(source)).toHaveLength(1);
    const numeric =
      "function number(value: number) { return value.toString(16).slice(-8); } function preview(value: string) { return value.toString(16).slice(0, 8); }";
    expect(findTruncations(numeric)).toHaveLength(1);
    const copied =
      "const value = 'text'; const copied = value; { const value = new Uint8Array(); const copied = value; copied.slice(0, 40); } copied.slice(0, 40);";
    expect(findTruncations(copied)).toHaveLength(1);
    expect(
      findTruncations(
        "function numeric(value: number) { const copied = value; return copied.toString(16).slice(-8); } function preview(value: string) { const copied = value; const cut = copied.slice; const alias = cut; return alias(0, 40); }",
      ),
    ).toHaveLength(1);
    expect(
      findTruncations(
        "const end = 40; function select(text: string, end: number) { return text.slice(0, end); } text.slice(0, end);",
      ),
    ).toHaveLength(1);
  });
  test("ASCII exceptions require a validator or an adjacent ASCII delimiter", () => {
    expect(
      findTruncations(
        'const HEX = /^[0-9a-f]{64}$/; HEX.test(value) ? value.slice(0, 32) : null; function prefix(name: string) { if (name[1] === ":") return name[0]; }',
      ),
    ).toEqual([]);
    expect(
      findTruncations(
        "const ANY = /^.{64}$/u; ANY.test(value) ? value.slice(0, 32) : null;",
      ).length,
    ).toBeGreaterThan(0);
    expect(
      findTruncations("function prefix(name: string) { return name[0]; }")
        .length,
    ).toBeGreaterThan(0);
    expect(
      findTruncations(
        'function prefix(name: string) { if (name[1] === ":") { const name = "👩"; return name[0]; } }',
      ).length,
    ).toBeGreaterThan(0);
    expect(
      findTruncations(
        'function id(value: number) { return value.toString(16).padStart(8, "0").slice(-8); } Math.random().toString(36).slice(2, 10);',
      ),
    ).toEqual([]);
    expect(
      findTruncations(
        "function text(value: string) { return value.toString().slice(0, 40); }",
      ).length,
    ).toBeGreaterThan(0);
  });
  test("the font ASCII exception is tied to the actual admitted-character filter", async () => {
    const file = "core/src/fonts/sfnt/parse.ts";
    const source = await Bun.file(path.join(PACKAGES, file)).text();
    expect(violationsFor(file, source)).toEqual([]);
    expect(
      violationsFor(file, source.replace("code < 0x7f", "code < 0x10ffff"))
        .length,
    ).toBeGreaterThan(0);
  });
  test("production bounded text passes through its owner", async () => {
    const offenders: string[] = [];
    for (const file of SOURCE_FILES) {
      for (const { line, reason } of violationsFor(
        file,
        await Bun.file(path.join(PACKAGES, file)).text(),
      ))
        offenders.push(`${file}:${line}: ${reason}`);
    }
    expect(offenders).toEqual([]);
  }, 10_000);
});
