import { describe, expect, test } from "bun:test";
import path from "node:path";
import ts from "typescript";
import { readFileSync } from "node:fs";

const PACKAGES = path.resolve(import.meta.dir, "../packages");
const OWNER = "utf8-boundaries/src/sliceText.ts";
const EXTENSIONS = ["ts", "tsx", "js", "jsx", "mts", "cts", "mjs", "cjs", "vue"];
const SOURCE_GLOB = `*/src/**/*.{${EXTENSIONS.join(",")}}`;
const CAP_NAME = /max|limit|radius|preview|snippet|context|truncate|unchanged.*edge/iu;
const TEXT_NAME =
  /^(?:text|message|name|value|json|label|description|title|raw|query|author|snippet|context)$/iu;

type Violation = { line: number; reason: string };

const SPACE_INSPECTOR = "core/src/layout-engine/measure/paragraphMeasureShared.ts";
const verifiedSpaceInspector = (text: string): boolean => {
  const source = ts.createSourceFile(SPACE_INSPECTOR, text, ts.ScriptTarget.Latest, true);
  const declaration = source.statements.find(
    (node) => ts.isFunctionDeclaration(node) && node.name?.text === "isSpaceOrTab",
  );
  if (
    !declaration ||
    !ts.isFunctionDeclaration(declaration) ||
    declaration.parameters.length !== 1 ||
    declaration.body?.statements.length !== 1
  )
    return false;
  const parameter = declaration.parameters.at(0)?.name;
  const statement = declaration.body.statements.at(0);
  if (
    !parameter ||
    !ts.isIdentifier(parameter) ||
    !statement ||
    !ts.isReturnStatement(statement) ||
    !statement.expression ||
    !ts.isBinaryExpression(statement.expression) ||
    statement.expression.operatorToken.kind !== ts.SyntaxKind.BarBarToken
  )
    return false;
  const compares = (node: ts.Expression, expected: string): boolean =>
    ts.isBinaryExpression(node) &&
    node.operatorToken.kind === ts.SyntaxKind.EqualsEqualsEqualsToken &&
    ts.isIdentifier(node.left) &&
    node.left.text === parameter.text &&
    ts.isStringLiteral(node.right) &&
    node.right.text === expected;
  return compares(statement.expression.left, " ") && compares(statement.expression.right, "\t");
};

const member = (node: ts.Expression): string | undefined => {
  if (ts.isPropertyAccessExpression(node)) return node.name.text;
  if (
    ts.isElementAccessExpression(node) &&
    node.argumentExpression &&
    ts.isStringLiteral(node.argumentExpression)
  )
    return node.argumentExpression.text;
  return undefined;
};

const receiver = (node: ts.Expression): ts.Expression | undefined =>
  ts.isPropertyAccessExpression(node) || ts.isElementAccessExpression(node)
    ? node.expression
    : undefined;

const findTruncations = (input: string, file = "fixture.ts"): Violation[] => {
  const text = file.endsWith(".vue")
    ? [...input.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/gu)]
        .map((match) => match[1] ?? "")
        .join("\n")
    : input;
  const source = ts.createSourceFile(
    file,
    text,
    ts.ScriptTarget.Latest,
    true,
    file.endsWith(".tsx") || file.endsWith(".jsx") ? ts.ScriptKind.TSX : ts.ScriptKind.TS,
  );
  const nodes: ts.Node[] = [];
  const collect = (node: ts.Node): void => {
    nodes.push(node);
    ts.forEachChild(node, collect);
  };
  collect(source);
  // Resolve identifiers to declarations rather than sharing facts by spelling.
  const bindings = new Map<ts.Node, Map<string, ts.Node>>();
  const isScope = (node: ts.Node): boolean =>
    ts.isSourceFile(node) ||
    ts.isBlock(node) ||
    ts.isFunctionLike(node) ||
    ts.isForStatement(node) ||
    ts.isForOfStatement(node) ||
    ts.isForInStatement(node) ||
    ts.isCatchClause(node) ||
    ts.isCaseBlock(node);
  const scopeOf = (node: ts.Node | undefined): ts.Node => {
    for (let current = node; current; current = current.parent)
      if (isScope(current)) return current;
    return source;
  };
  const bind = (name: ts.BindingName, declaration: ts.Node, scope: ts.Node): void => {
    if (!ts.isIdentifier(name)) {
      for (const element of name.elements)
        if (ts.isBindingElement(element)) bind(element.name, element, scope);
      return;
    }
    let names = bindings.get(scope);
    if (!names) {
      names = new Map();
      bindings.set(scope, names);
    }
    names.set(name.text, declaration);
  };
  for (const node of nodes) {
    if (ts.isParameter(node)) bind(node.name, node, scopeOf(node.parent));
    if (ts.isVariableDeclaration(node)) {
      let scope = scopeOf(node.parent);
      if (
        ts.isVariableDeclarationList(node.parent) &&
        !(node.parent.flags & ts.NodeFlags.BlockScoped)
      ) {
        for (let current: ts.Node | undefined = node.parent; current; current = current.parent) {
          if (ts.isFunctionLike(current) || ts.isSourceFile(current)) {
            scope = current;
            break;
          }
        }
      }
      bind(node.name, node, scope);
    }
    if (ts.isFunctionDeclaration(node) && node.name) bind(node.name, node, scopeOf(node.parent));
    if (ts.isImportSpecifier(node)) bind(node.name, node, scopeOf(node.parent));
    if (ts.isImportClause(node) && node.name) bind(node.name, node, scopeOf(node.parent));
    if (ts.isNamespaceImport(node)) bind(node.name, node, scopeOf(node.parent));
  }
  const bindingOf = (identifier: ts.Identifier): ts.Node => {
    for (
      let scope: ts.Node | undefined = scopeOf(identifier.parent);
      scope;
      scope = scope === source ? undefined : scopeOf(scope.parent)
    ) {
      const declaration = bindings.get(scope)?.get(identifier.text);
      if (declaration) return declaration;
    }
    return identifier;
  };
  const strings = new Set<ts.Node>();
  const arrays = new Set<ts.Node>();
  const stringArrays = new Set<ts.Node>();
  const stringFunctions = new Set<ts.Node>();
  const methods = new Map<ts.Node, { receiver: ts.Expression; method: string }>();
  const caps = new Set<ts.Node>();
  const numbers = new Set<ts.Node>();
  const asciiValidators = new Set<ts.Node>();
  const regexes = new Set<ts.Node>();
  const inspectors = new Set<ts.Node>();
  const mutated = new Set<ts.Node>();
  for (const node of nodes) {
    if (
      (ts.isPrefixUnaryExpression(node) || ts.isPostfixUnaryExpression(node)) &&
      (node.operator === ts.SyntaxKind.PlusPlusToken ||
        node.operator === ts.SyntaxKind.MinusMinusToken) &&
      ts.isIdentifier(node.operand)
    )
      mutated.add(bindingOf(node.operand));
    if (
      ts.isBinaryExpression(node) &&
      node.operatorToken.kind >= ts.SyntaxKind.FirstAssignment &&
      node.operatorToken.kind <= ts.SyntaxKind.LastAssignment &&
      ts.isIdentifier(node.left)
    )
      mutated.add(bindingOf(node.left));
  }
  const numericCall = (node: ts.Expression): boolean => {
    if (!ts.isCallExpression(node)) return false;
    if (ts.isIdentifier(node.expression))
      return (
        /^(?:parseInt|parseFloat|Number)$/u.test(node.expression.text) &&
        bindingOf(node.expression) === node.expression
      );
    const owner = receiver(node.expression);
    return (
      owner !== undefined &&
      ts.isIdentifier(owner) &&
      owner.text === "Number" &&
      bindingOf(owner) === owner &&
      /^(?:parseInt|parseFloat)$/u.test(member(node.expression) ?? "")
    );
  };
  for (const node of nodes) {
    if ((ts.isParameter(node) || ts.isVariableDeclaration(node)) && ts.isIdentifier(node.name)) {
      if (node.type?.kind === ts.SyntaxKind.StringKeyword) strings.add(node);
      if (node.type?.kind === ts.SyntaxKind.NumberKeyword) numbers.add(node);
      if (node.type?.getText(source) === "RegExp") regexes.add(node);
      if (/^RegExp(?:Match|Exec)Array$/u.test(node.type?.getText(source) ?? "")) arrays.add(node);
      if (
        node.type &&
        (ts.isArrayTypeNode(node.type) ||
          /^(?:readonly\s+)?(?:Uint\d+Array|Int\d+Array|ArrayBuffer|Buffer|Array<|ReadonlyArray<)/u.test(
            node.type.getText(source),
          ))
      )
        arrays.add(node);
      if (
        node.type &&
        ts.isArrayTypeNode(node.type) &&
        node.type.elementType.kind === ts.SyntaxKind.StringKeyword
      )
        stringArrays.add(node);
    }
    if (
      ts.isFunctionDeclaration(node) &&
      node.name &&
      node.type?.kind === ts.SyntaxKind.StringKeyword
    )
      stringFunctions.add(node);
    if (
      ts.isVariableDeclaration(node) &&
      ts.isIdentifier(node.name) &&
      node.initializer &&
      ts.isArrowFunction(node.initializer) &&
      node.initializer.type?.kind === ts.SyntaxKind.StringKeyword
    )
      stringFunctions.add(node);
    if (
      ts.isVariableDeclaration(node) &&
      ts.isIdentifier(node.name) &&
      node.initializer?.kind === ts.SyntaxKind.RegularExpressionLiteral &&
      /^\/\^\[[\da-zA-Z-]+\]\{\d+\}\$\/[gu]*$/u.test(node.initializer.getText(source))
    )
      asciiValidators.add(node);
    if (
      ts.isVariableDeclaration(node) &&
      node.initializer &&
      (node.initializer.kind === ts.SyntaxKind.RegularExpressionLiteral ||
        (ts.isNewExpression(node.initializer) &&
          ts.isIdentifier(node.initializer.expression) &&
          node.initializer.expression.text === "RegExp" &&
          bindingOf(node.initializer.expression) === node.initializer.expression))
    )
      regexes.add(node);
    if (
      ts.isVariableDeclaration(node) &&
      node.initializer &&
      ts.isArrowFunction(node.initializer) &&
      !ts.isBlock(node.initializer.body) &&
      numericCall(node.initializer.body)
    )
      inspectors.add(node);
    if (
      ts.isImportSpecifier(node) &&
      (node.propertyName?.text ?? node.name.text) === "isSpaceOrTab"
    ) {
      const declaration = node.parent.parent.parent;
      if (
        ts.isImportDeclaration(declaration) &&
        ts.isStringLiteral(declaration.moduleSpecifier) &&
        path.resolve(PACKAGES, path.dirname(file), `${declaration.moduleSpecifier.text}.ts`) ===
          path.join(PACKAGES, SPACE_INSPECTOR) &&
        verifiedSpaceInspector(readFileSync(path.join(PACKAGES, SPACE_INSPECTOR), "utf8"))
      )
        inspectors.add(node);
    }
  }
  const isArray = (node: ts.Expression): boolean => {
    if (ts.isArrayLiteralExpression(node)) return true;
    if (ts.isIdentifier(node)) return arrays.has(bindingOf(node));
    if (ts.isNewExpression(node))
      return /^(?:Uint\d+Array|Int\d+Array|ArrayBuffer|Buffer|Array)$/u.test(
        node.expression.getText(source),
      );
    if (ts.isCallExpression(node)) {
      if (
        member(node.expression) === "split" ||
        member(node.expression) === "map" ||
        member(node.expression) === "filter"
      )
        return true;
      if (
        member(node.expression) === "from" &&
        receiver(node.expression)?.getText(source) === "Array"
      )
        return true;
      const owner = receiver(node.expression);
      if (
        member(node.expression) === "exec" &&
        owner &&
        (owner.kind === ts.SyntaxKind.RegularExpressionLiteral ||
          (ts.isIdentifier(owner) && regexes.has(bindingOf(owner))))
      )
        return true;
      if (member(node.expression) === "match" && owner && isString(owner)) return true;
    }
    return false;
  };
  const isStringArray = (node: ts.Expression): boolean => {
    if (ts.isIdentifier(node)) return stringArrays.has(bindingOf(node));
    if (ts.isArrayLiteralExpression(node))
      return node.elements.every((element) => isString(element));
    if (!ts.isCallExpression(node)) return false;
    const owner = receiver(node.expression);
    if (!owner) return false;
    if (member(node.expression) === "split") return isString(owner);
    return member(node.expression) === "filter" && isStringArray(owner);
  };
  const isString = (node: ts.Expression): boolean => {
    if (isArray(node)) return false;
    if (
      ts.isStringLiteral(node) ||
      ts.isNoSubstitutionTemplateLiteral(node) ||
      ts.isTemplateExpression(node)
    )
      return true;
    if (ts.isIdentifier(node))
      return (
        !numbers.has(bindingOf(node)) && (strings.has(bindingOf(node)) || TEXT_NAME.test(node.text))
      );
    if (ts.isPropertyAccessExpression(node) || ts.isElementAccessExpression(node))
      return /^(?:text|message|name|title|description|label)$/u.test(member(node) ?? "");
    if (ts.isCallExpression(node)) {
      if (ts.isIdentifier(node.expression))
        return node.expression.text === "String" || stringFunctions.has(bindingOf(node.expression));
      const method = member(node.expression);
      if (method === "join") return true;
      if (method === "toString") return true;
      const owner = receiver(node.expression);
      return (
        owner !== undefined &&
        /^(?:slice|substring|substr|trim|trimStart|trimEnd|replace|replaceAll|toUpperCase|toLowerCase|padStart|padEnd|toString)$/u.test(
          method ?? "",
        ) &&
        isString(owner)
      );
    }
    if (ts.isBinaryExpression(node) && node.operatorToken.kind === ts.SyntaxKind.PlusToken)
      return isString(node.left) || isString(node.right);
    return false;
  };
  let changed = true;
  const coordinateWalk = (declaration: ts.VariableDeclaration): boolean => {
    if (
      ts.isVariableDeclarationList(declaration.parent) &&
      ts.isForStatement(declaration.parent.parent)
    )
      return true;
    return nodes.some((node) => {
      if (!ts.isWhileStatement(node) && !ts.isForStatement(node)) return false;
      const condition = ts.isWhileStatement(node) ? node.expression : node.condition;
      if (!condition) return false;
      let indexed = false;
      const inspect = (part: ts.Node): void => {
        if (
          ts.isElementAccessExpression(part) &&
          part.argumentExpression &&
          ts.isIdentifier(part.argumentExpression) &&
          bindingOf(part.argumentExpression) === declaration &&
          isString(part.expression)
        )
          indexed = true;
        ts.forEachChild(part, inspect);
      };
      inspect(condition);
      return indexed;
    });
  };
  while (changed) {
    changed = false;
    for (const node of nodes) {
      if (
        ts.isVariableDeclaration(node) &&
        ts.isIdentifier(node.name) &&
        ts.isVariableDeclarationList(node.parent) &&
        ts.isForOfStatement(node.parent.parent)
      ) {
        const iterable = node.parent.parent.expression;
        const owner = ts.isCallExpression(iterable) ? receiver(iterable.expression) : undefined;
        if (
          ts.isCallExpression(iterable) &&
          member(iterable.expression) === "matchAll" &&
          owner &&
          isString(owner) &&
          !arrays.has(node)
        ) {
          arrays.add(node);
          changed = true;
        }
      }
      if (
        ts.isCallExpression(node) &&
        /^(?:map|filter|forEach)$/u.test(member(node.expression) ?? "")
      ) {
        const owner = receiver(node.expression);
        const callback = node.arguments.at(0);
        if (
          owner &&
          isStringArray(owner) &&
          callback &&
          (ts.isArrowFunction(callback) || ts.isFunctionExpression(callback))
        ) {
          const parameter = callback.parameters.at(0)?.name;
          if (parameter && ts.isIdentifier(parameter) && !strings.has(bindingOf(parameter))) {
            strings.add(bindingOf(parameter));
            changed = true;
          }
        }
      }
      if (!ts.isVariableDeclaration(node) || !node.initializer) continue;
      if (ts.isObjectBindingPattern(node.name) && isString(node.initializer)) {
        for (const element of node.name.elements) {
          if (!ts.isIdentifier(element.name)) continue;
          const method =
            element.propertyName?.getText(source).replace(/["']/gu, "") ?? element.name.text;
          if (
            /^(?:slice|substring|substr)$/u.test(method) &&
            !methods.has(bindingOf(element.name))
          ) {
            methods.set(bindingOf(element.name), { receiver: node.initializer, method });
            changed = true;
          }
        }
      }
      if (!ts.isIdentifier(node.name)) continue;
      if (
        ts.isIdentifier(node.initializer) &&
        regexes.has(bindingOf(node.initializer)) &&
        !regexes.has(node)
      ) {
        regexes.add(node);
        changed = true;
      }
      if (
        ts.isIdentifier(node.initializer) &&
        inspectors.has(bindingOf(node.initializer)) &&
        !inspectors.has(node)
      ) {
        inspectors.add(node);
        changed = true;
      }
      if (
        (ts.isNumericLiteral(node.initializer) ||
          (ts.isIdentifier(node.initializer) && numbers.has(bindingOf(node.initializer)))) &&
        !numbers.has(node)
      ) {
        numbers.add(node);
        changed = true;
      }
      if (isArray(node.initializer) && !arrays.has(bindingOf(node.name))) {
        arrays.add(node);
        changed = true;
      }
      if (isStringArray(node.initializer) && !stringArrays.has(bindingOf(node.name))) {
        stringArrays.add(node);
        changed = true;
      }
      if (isString(node.initializer) && !strings.has(bindingOf(node.name))) {
        strings.add(node);
        changed = true;
      }
      if (
        ts.isVariableDeclarationList(node.parent) &&
        (!mutated.has(node) ||
          !/(?:offset|index|idx|start|end|cursor|position|^i$|^j$)/iu.test(node.name.text) ||
          !coordinateWalk(node)) &&
        ((ts.isNumericLiteral(node.initializer) && Number(node.initializer.text) >= 1) ||
          (ts.isIdentifier(node.initializer) && caps.has(bindingOf(node.initializer))))
      ) {
        if (!caps.has(bindingOf(node.name))) {
          caps.add(bindingOf(node.name));
          changed = true;
        }
      }
      const owner = receiver(node.initializer);
      const method = member(node.initializer);
      if (
        owner &&
        isString(owner) &&
        /^(?:slice|substring|substr)$/u.test(method ?? "") &&
        !methods.has(bindingOf(node.name))
      ) {
        methods.set(bindingOf(node.name), { receiver: owner, method: method ?? "" });
        changed = true;
      }
      if (
        ts.isIdentifier(node.initializer) &&
        methods.has(bindingOf(node.initializer)) &&
        !methods.has(bindingOf(node.name))
      ) {
        methods.set(bindingOf(node.name), methods.get(bindingOf(node.initializer))!);
        changed = true;
      }
      if (ts.isCallExpression(node.initializer) && member(node.initializer.expression) === "bind") {
        const reference = receiver(node.initializer.expression);
        const boundOwner = reference ? receiver(reference) : undefined;
        const boundMethod = reference ? member(reference) : undefined;
        if (
          boundOwner &&
          isString(boundOwner) &&
          /^(?:slice|substring|substr)$/u.test(boundMethod ?? "") &&
          !methods.has(bindingOf(node.name))
        ) {
          methods.set(bindingOf(node.name), { receiver: boundOwner, method: boundMethod ?? "" });
          changed = true;
        }
      }
    }
  }
  const sameBinding = (left: ts.Identifier, right: ts.Identifier): boolean => {
    const first = bindingOf(left);
    const second = bindingOf(right);
    return first === second || (first === left && second === right && left.text === right.text);
  };
  const coordinate = (node: ts.Expression): boolean => {
    if (ts.isPropertyAccessExpression(node))
      return /^(?:index|offset|position|start|end)$/iu.test(node.name.text);
    if (!ts.isIdentifier(node) || CAP_NAME.test(node.text) || caps.has(bindingOf(node)))
      return false;
    if (/(?:offset|index|idx|start|end|cursor|position|break|^at$)/iu.test(node.text)) return true;
    const declaration = bindingOf(node);
    return (
      ts.isVariableDeclaration(declaration) &&
      declaration.initializer !== undefined &&
      ts.isCallExpression(declaration.initializer) &&
      /^(?:indexOf|lastIndexOf|search)$/u.test(member(declaration.initializer.expression) ?? "")
    );
  };
  const cap = (node: ts.Expression | undefined): boolean => {
    if (!node) return false;
    if (ts.isNumericLiteral(node)) return Number(node.text) >= 1;
    if (ts.isIdentifier(node)) return caps.has(bindingOf(node)) || CAP_NAME.test(node.text);
    if (ts.isPrefixUnaryExpression(node) && ts.isNumericLiteral(node.operand))
      return Number(node.operand.text) >= 1;
    if (ts.isCallExpression(node)) return node.arguments.some((argument) => cap(argument));
    if (ts.isBinaryExpression(node)) {
      // Small offsets beside parser/selection coordinates retain UTF-16 units.
      if (
        (coordinate(node.left) &&
          ts.isNumericLiteral(node.right) &&
          Number(node.right.text) <= 4) ||
        (coordinate(node.right) && ts.isNumericLiteral(node.left) && Number(node.left.text) <= 4)
      )
        return false;
      return cap(node.left) || cap(node.right);
    }
    if (ts.isPropertyAccessExpression(node)) return CAP_NAME.test(node.name.text);
    return false;
  };
  const isInspection = (node: ts.Node): boolean => {
    // Numeric and regexp predicates inspect units; they do not emit clipped text.
    let parent = node.parent;
    while (ts.isNonNullExpression(parent) || ts.isParenthesizedExpression(parent))
      parent = parent.parent;
    if (!ts.isCallExpression(parent)) return false;
    if (numericCall(parent)) return true;
    if (ts.isIdentifier(parent.expression) && inspectors.has(bindingOf(parent.expression)))
      return true;
    const owner = receiver(parent.expression);
    return (
      member(parent.expression) === "test" &&
      owner !== undefined &&
      (owner.kind === ts.SyntaxKind.RegularExpressionLiteral ||
        (ts.isIdentifier(owner) && regexes.has(bindingOf(owner))))
    );
  };
  const asciiTailCondition = (
    condition: ts.Expression,
    owner: ts.Identifier,
    amount: number,
    seen = new Set<ts.Node>(),
  ): boolean => {
    if (seen.has(condition)) return false;
    seen.add(condition);
    if (ts.isParenthesizedExpression(condition))
      return asciiTailCondition(condition.expression, owner, amount, seen);
    if (ts.isIdentifier(condition)) {
      const declaration = bindingOf(condition);
      return (
        ts.isVariableDeclaration(declaration) &&
        declaration.initializer !== undefined &&
        asciiTailCondition(declaration.initializer, owner, amount, seen)
      );
    }
    if (ts.isBinaryExpression(condition)) {
      if (condition.operatorToken.kind === ts.SyntaxKind.AmpersandAmpersandToken)
        return (
          asciiTailCondition(condition.left, owner, amount, seen) ||
          asciiTailCondition(condition.right, owner, amount, seen)
        );
      if (condition.operatorToken.kind === ts.SyntaxKind.BarBarToken)
        return (
          asciiTailCondition(condition.left, owner, amount, new Set(seen)) &&
          asciiTailCondition(condition.right, owner, amount, new Set(seen))
        );
    }
    if (!ts.isCallExpression(condition) || member(condition.expression) !== "endsWith")
      return false;
    const tested = receiver(condition.expression);
    const suffix = condition.arguments.at(0);
    return (
      tested !== undefined &&
      ts.isIdentifier(tested) &&
      sameBinding(tested, owner) &&
      suffix !== undefined &&
      ts.isStringLiteral(suffix) &&
      suffix.text.length === amount &&
      suffix.text.length > 0 &&
      Array.from(suffix.text).every((character) => character.charCodeAt(0) <= 0x7f)
    );
  };
  const protectedAsciiTail = (node: ts.CallExpression, owner: ts.Expression): boolean => {
    if (!ts.isIdentifier(owner)) return false;
    const end =
      node.arguments.at(1) ?? (node.arguments.length === 1 ? node.arguments.at(0) : undefined);
    let amount: number | undefined;
    if (
      end &&
      ts.isPrefixUnaryExpression(end) &&
      end.operator === ts.SyntaxKind.MinusToken &&
      ts.isNumericLiteral(end.operand)
    )
      amount = Number(end.operand.text);
    if (
      end &&
      ts.isBinaryExpression(end) &&
      end.operatorToken.kind === ts.SyntaxKind.MinusToken &&
      ts.isPropertyAccessExpression(end.left) &&
      end.left.name.text === "length" &&
      ts.isIdentifier(end.left.expression) &&
      sameBinding(end.left.expression, owner) &&
      ts.isNumericLiteral(end.right)
    )
      amount = Number(end.right.text);
    if (amount === undefined) return false;
    for (let parent = node.parent; parent; parent = parent.parent) {
      if (ts.isFunctionLike(parent)) break;
      if (
        ts.isIfStatement(parent) &&
        node.pos >= parent.thenStatement.pos &&
        node.end <= parent.thenStatement.end &&
        asciiTailCondition(parent.expression, owner, amount)
      )
        return true;
      if (
        ts.isConditionalExpression(parent) &&
        node.pos >= parent.whenTrue.pos &&
        node.end <= parent.whenTrue.end &&
        asciiTailCondition(parent.condition, owner, amount)
      )
        return true;
    }
    return false;
  };
  const asciiNumber = (node: ts.Expression): boolean => {
    if (!ts.isCallExpression(node)) return false;
    const method = member(node.expression);
    const owner = receiver(node.expression);
    if (!owner) return false;
    if (/^(?:toUpperCase|toLowerCase|padStart|padEnd)$/u.test(method ?? "")) {
      const padding = node.arguments.at(1);
      return (
        asciiNumber(owner) &&
        (!padding ||
          (ts.isStringLiteral(padding) &&
            Array.from(padding.text).every((character) => character.charCodeAt(0) <= 0x7f)))
      );
    }
    if (
      method !== "toString" ||
      !/^(?:8|10|16|36)$/u.test(node.arguments.at(0)?.getText(source) ?? "")
    )
      return false;
    return (
      ts.isNumericLiteral(owner) ||
      (ts.isIdentifier(owner) && numbers.has(bindingOf(owner))) ||
      (ts.isCallExpression(owner) &&
        ts.isPropertyAccessExpression(owner.expression) &&
        ts.isIdentifier(owner.expression.expression) &&
        bindingOf(owner.expression.expression) === owner.expression.expression &&
        /^(?:Math\.random|Date\.now)$/u.test(owner.expression.getText(source)))
    );
  };
  const violations: Violation[] = [];
  const add = (node: ts.Node, reason: string): void => {
    violations.push({
      line: source.getLineAndCharacterOfPosition(node.getStart(source)).line + 1,
      reason,
    });
  };
  for (const node of nodes) {
    if (
      ts.isElementAccessExpression(node) &&
      node.argumentExpression &&
      ts.isNumericLiteral(node.argumentExpression) &&
      node.argumentExpression.text === "0" &&
      isString(node.expression)
    ) {
      // A following ASCII delimiter proves the first unit is not half a pair.
      let delimiterGuard = false;
      for (let parent = node.parent; parent; parent = parent.parent) {
        if (!ts.isIfStatement(parent)) continue;
        const condition = parent.expression;
        if (
          ts.isBinaryExpression(condition) &&
          condition.operatorToken.kind === ts.SyntaxKind.EqualsEqualsEqualsToken &&
          ts.isElementAccessExpression(condition.left) &&
          condition.left.argumentExpression?.getText(source) === "1" &&
          ts.isStringLiteral(condition.right) &&
          condition.right.text === ":" &&
          ts.isIdentifier(node.expression) &&
          ts.isIdentifier(condition.left.expression) &&
          sameBinding(node.expression, condition.left.expression) &&
          node.pos >= parent.thenStatement.pos &&
          node.end <= parent.thenStatement.end
        )
          delimiterGuard = true;
        break;
      }
      if (!delimiterGuard && !isInspection(node))
        add(node, "string initials must use grapheme slicing");
    }
    if (!ts.isCallExpression(node)) continue;
    const alias = ts.isIdentifier(node.expression)
      ? methods.get(bindingOf(node.expression))
      : undefined;
    const owner = alias?.receiver ?? receiver(node.expression);
    const method = alias?.method ?? member(node.expression);
    const first = node.arguments.at(0);
    if (
      owner &&
      isString(owner) &&
      (method === "at" || method === "charAt") &&
      ((first && ts.isNumericLiteral(first) && Number(first.text) === 0) ||
        (method === "charAt" && first === undefined))
    ) {
      if (!asciiNumber(owner) && !isInspection(node))
        add(node, "string initials must use grapheme slicing");
      continue;
    }
    if (!owner || !isString(owner) || !/^(?:slice|substring|substr)$/u.test(method ?? "")) continue;
    if (asciiNumber(owner)) continue;
    if (isInspection(node) || protectedAsciiTail(node, owner)) continue;
    // Parsing a hexadecimal numeric field does not emit the sliced text.
    const parent = node.parent;
    const numericParser =
      ts.isCallExpression(parent) &&
      ((ts.isIdentifier(parent.expression) &&
        parent.expression.text === "parseInt" &&
        bindingOf(parent.expression) === parent.expression) ||
        (ts.isPropertyAccessExpression(parent.expression) &&
          parent.expression.name.text === "parseInt" &&
          ts.isIdentifier(parent.expression.expression) &&
          parent.expression.expression.text === "Number" &&
          bindingOf(parent.expression.expression) === parent.expression.expression));
    if (
      ts.isCallExpression(parent) &&
      numericParser &&
      parent.arguments.at(1)?.getText(source) === "16"
    )
      continue;
    if (
      ts.isConditionalExpression(parent) &&
      parent.whenTrue === node &&
      ts.isCallExpression(parent.condition) &&
      member(parent.condition.expression) === "test"
    ) {
      const validator = receiver(parent.condition.expression);
      const validated = parent.condition.arguments.at(0);
      if (
        validator &&
        ts.isIdentifier(validator) &&
        asciiValidators.has(bindingOf(validator)) &&
        validated &&
        ts.isIdentifier(validated) &&
        ts.isIdentifier(owner) &&
        sameBinding(validated, owner)
      )
        continue;
    }
    if (
      method === "substr" ||
      cap(node.arguments.at(1)) ||
      (node.arguments.length === 1 &&
        ts.isPrefixUnaryExpression(node.arguments[0]!) &&
        cap(node.arguments[0]))
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
    const source = ts.createSourceFile(file, text, ts.ScriptTarget.Latest, true);
    const ascii = source.statements.find(
      (node) =>
        ts.isVariableStatement(node) &&
        node.declarationList.declarations.some(
          ({ name }) => ts.isIdentifier(name) && name.text === "sanitizePostScriptName",
        ),
    );
    if (ascii?.getText(source).replace(/\s+/gu, "") === ASCII_FONT_FUNCTION)
      return findTruncations(
        text.slice(0, ascii.getStart(source)) +
          " ".repeat(ascii.end - ascii.getStart(source)) +
          text.slice(ascii.end),
        file,
      );
  }
  return findTruncations(text, file);
};

describe("text truncation boundaries", () => {
  test("every production package and source extension rejects bounded text and initials", async () => {
    const packages = new Set<string>();
    for await (const file of new Bun.Glob(SOURCE_GLOB).scan({ cwd: PACKAGES }))
      packages.add(file.split("/")[0]!);
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
          const text = extension === "vue" ? `<script setup>${fixture}</script>` : fixture;
          expect({ file, fixture, rejected: violationsFor(file, text).length > 0 }).toEqual({
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
    expect(verifiedSpaceInspector(helper.replace('char === " "', 'char === "👩"'))).toBe(false);
    expect(
      findTruncations(
        'import { isSpaceOrTab } from "./paragraphMeasureShared"; isSpaceOrTab(text.at(0)); isSpaceOrTab(text[0]);',
        "core/src/layout-engine/measure/fixture.ts",
      ),
    ).toEqual([]);
    expect(
      findTruncations("const isSpaceOrTab = (text: string) => text; isSpaceOrTab(text[0]);"),
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
      findTruncations("const ANY = /^.{64}$/u; ANY.test(value) ? value.slice(0, 32) : null;")
        .length,
    ).toBeGreaterThan(0);
    expect(
      findTruncations("function prefix(name: string) { return name[0]; }").length,
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
      findTruncations("function text(value: string) { return value.toString().slice(0, 40); }")
        .length,
    ).toBeGreaterThan(0);
  });
  test("the font ASCII exception is tied to the actual admitted-character filter", async () => {
    const file = "core/src/fonts/sfnt/parse.ts";
    const source = await Bun.file(path.join(PACKAGES, file)).text();
    expect(violationsFor(file, source)).toEqual([]);
    expect(
      violationsFor(file, source.replace("code < 0x7f", "code < 0x10ffff")).length,
    ).toBeGreaterThan(0);
  });
  test("production bounded text passes through its owner", async () => {
    const offenders: string[] = [];
    for await (const file of new Bun.Glob(SOURCE_GLOB).scan({ cwd: PACKAGES })) {
      for (const { line, reason } of violationsFor(
        file,
        await Bun.file(path.join(PACKAGES, file)).text(),
      ))
        offenders.push(`${file}:${line}: ${reason}`);
    }
    expect(offenders).toEqual([]);
  }, 10_000);
});
