/** DOM scroll writes belong to the shared editor scroll-root boundary. */
import { describe, expect, test } from "bun:test";
import path from "node:path";
import ts from "typescript";

const ROOT = path.resolve(import.meta.dir, "..");
const SCROLL_WRITER = "packages/core/src/paged-layout/editorScrollRoot.ts";
const OUTLINE_LISTS = {
  "packages/react/src/ui/defaults/outline-rail.tsx": "listRef.current",
  "packages/vue/src/components/ui/OutlineRail.vue": "listRef.value",
} as const satisfies Record<string, string>;
const METHODS = new Set(["scrollIntoView", "scrollTo", "scrollBy"]);
const COORDINATES = new Set(["scrollTop", "scrollLeft"]);

const parseSource = (file: string, text: string) => {
  // Keep offsets intact so Vue diagnostics point to the original source line.
  const script = file.endsWith(".vue")
    ? text.replace(/<script\b[^>]*>([\s\S]*?)<\/script>|[^<]+|</gu, (match, body: unknown) =>
        typeof body === "string"
          ? match.slice(0, match.indexOf(body)).replace(/[^\n]/gu, " ") +
            body +
            match.slice(match.indexOf(body) + body.length).replace(/[^\n]/gu, " ")
          : match.replace(/[^\n]/gu, " "),
      )
    : text;
  return ts.createSourceFile(
    file,
    script,
    ts.ScriptTarget.Latest,
    true,
    file.endsWith(".tsx") ? ts.ScriptKind.TSX : ts.ScriptKind.TS,
  );
};

const propertyName = (node: ts.Node): string | undefined => {
  if (ts.isPropertyAccessExpression(node)) return node.name.text;
  if (ts.isElementAccessExpression(node) && ts.isStringLiteral(node.argumentExpression)) {
    return node.argumentExpression.text;
  }
  return undefined;
};

const receiver = (node: ts.Node): ts.Expression | undefined =>
  ts.isPropertyAccessExpression(node) || ts.isElementAccessExpression(node)
    ? node.expression
    : undefined;

const scopeOf = (node: ts.Node): ts.Node => {
  let scope = node.parent;
  while (!ts.isSourceFile(scope) && !ts.isBlock(scope) && !ts.isFunctionLike(scope)) {
    scope = scope.parent;
  }
  return scope;
};

const ancestors = (node: ts.Node): ts.Node[] => {
  const result: ts.Node[] = [];
  let current: ts.Node | undefined = node;
  while (current) {
    result.push(current);
    current = current.parent;
  }
  return result;
};

type ScrollScanOptions = {
  file: string;
  text: string;
  transactionFunctions?: ReadonlySet<string>;
};

const scrollViolations = ({ file, text, transactionFunctions = new Set() }: ScrollScanOptions) => {
  if (file === SCROLL_WRITER) return { violations: [], outlineWrites: 0 };
  const source = parseSource(file, text);
  const bindings: Array<ts.VariableDeclaration | ts.ParameterDeclaration | ts.BindingElement> = [];
  const collect = (node: ts.Node): void => {
    if (ts.isVariableDeclaration(node) || ts.isParameter(node) || ts.isBindingElement(node))
      bindings.push(node);
    ts.forEachChild(node, collect);
  };
  collect(source);

  const bindingOf = (node: ts.Identifier) => {
    const scopes = ancestors(node);
    return bindings
      .filter(
        (candidate) =>
          ts.isIdentifier(candidate.name) &&
          candidate.name.text === node.text &&
          candidate.pos < node.pos &&
          scopes.includes(scopeOf(candidate)),
      )
      .sort(
        (left, right) =>
          scopes.indexOf(scopeOf(left)) - scopes.indexOf(scopeOf(right)) || right.pos - left.pos,
      )
      .at(0);
  };
  const hasTransactionReturn = (node: ts.Node): boolean => {
    if (ts.isFunctionTypeNode(node)) return node.type.getText(source) === "Transaction";
    return ts.forEachChild(node, (child) => hasTransactionReturn(child) || undefined) ?? false;
  };
  const isTransaction = (node: ts.Expression, seen = new Set<ts.Node>()): boolean => {
    if (seen.has(node)) return false;
    seen.add(node);
    if (ts.isParenthesizedExpression(node) || ts.isNonNullExpression(node)) {
      return isTransaction(node.expression, seen);
    }
    if (
      ts.isBinaryExpression(node) &&
      node.operatorToken.kind === ts.SyntaxKind.QuestionQuestionToken
    ) {
      return isTransaction(node.left, seen) && isTransaction(node.right, seen);
    }
    if (propertyName(node) === "tr") return true;
    if (ts.isCallExpression(node)) {
      const object = receiver(node.expression);
      if (object) return isTransaction(object, seen);
      if (!ts.isIdentifier(node.expression)) return false;
      if (transactionFunctions.has(node.expression.text)) return true;
      const binding = bindingOf(node.expression);
      return (
        binding !== undefined &&
        !ts.isBindingElement(binding) &&
        binding.type !== undefined &&
        hasTransactionReturn(binding.type)
      );
    }
    if (!ts.isIdentifier(node)) return false;
    const binding = bindingOf(node);
    if (!binding) return false;
    if (ts.isBindingElement(binding)) {
      return (binding.propertyName?.getText(source) ?? binding.name.getText(source)) === "tr";
    }
    if (binding.type && /\bTransaction\b/u.test(binding.type.getText(source))) return true;
    return binding.initializer !== undefined && isTransaction(binding.initializer, seen);
  };

  const isOutlineListWrite = (node: ts.BinaryExpression): boolean => {
    if (!(file in OUTLINE_LISTS) || propertyName(node.left) !== "scrollTop") return false;
    const expectedRef = Object.entries(OUTLINE_LISTS)
      .find(([name]) => name === file)
      ?.at(1);
    const object = receiver(node.left);
    if (!object || object.getText(source) !== "list") return false;
    if (
      !["list.scrollTop = top", "list.scrollTop = bottom - list.clientHeight"].includes(
        node.getText(source),
      )
    ) {
      return false;
    }
    const callback = ancestors(node).find((ancestor) => ts.isArrowFunction(ancestor));
    if (!callback || !ts.isArrowFunction(callback) || !ts.isBlock(callback.body)) return false;
    const call = callback.parent;
    if (!ts.isCallExpression(call) || !ts.isIdentifier(call.expression)) return false;
    const isReact =
      call.expression.text === "useEffect" &&
      call.arguments.at(1)?.getText(source) === "[activeIndex, presentation]";
    const isVue =
      call.expression.text === "watch" &&
      call.arguments.at(0)?.getText(source) === "[activeIndex, () => props.presentation]";
    if (!isReact && !isVue) return false;
    return callback.body.statements.some(
      (statement) =>
        ts.isVariableStatement(statement) &&
        statement.declarationList.declarations.some(
          (declaration) =>
            ts.isIdentifier(declaration.name) &&
            declaration.name.text === "list" &&
            declaration.initializer?.getText(source) === expectedRef,
        ),
    );
  };

  const violations: string[] = [];
  let outlineWrites = 0;
  const report = (node: ts.Node): void => {
    const { line } = source.getLineAndCharacterOfPosition(node.getStart(source));
    violations.push(`${file}:${line + 1}: ${node.getText(source)}`);
  };
  const visit = (node: ts.Node): void => {
    if (ts.isCallExpression(node) && METHODS.has(propertyName(node.expression) ?? "")) {
      const object = receiver(node.expression);
      const transactionScroll =
        propertyName(node.expression) === "scrollIntoView" &&
        node.arguments.length === 0 &&
        object !== undefined &&
        isTransaction(object);
      if (!transactionScroll) report(node);
    }
    if (
      ts.isBinaryExpression(node) &&
      node.operatorToken.kind >= ts.SyntaxKind.FirstAssignment &&
      node.operatorToken.kind <= ts.SyntaxKind.LastAssignment &&
      COORDINATES.has(propertyName(node.left) ?? "")
    ) {
      if (isOutlineListWrite(node)) outlineWrites += 1;
      else report(node);
    }
    if (
      (ts.isPrefixUnaryExpression(node) || ts.isPostfixUnaryExpression(node)) &&
      (node.operator === ts.SyntaxKind.PlusPlusToken ||
        node.operator === ts.SyntaxKind.MinusMinusToken) &&
      COORDINATES.has(propertyName(node.operand) ?? "")
    )
      report(node);
    ts.forEachChild(node, visit);
  };
  visit(source);
  return { violations, outlineWrites };
};

const isProductionSource = (file: string): boolean =>
  !/^packages\/playground(?:-vue)?\//u.test(file) &&
  !/(?:^|\/)(?:__tests__|__fixtures__|test|tests)\//u.test(file) &&
  !/\.(?:test|spec|typecheck)\.[^.]+$/u.test(file);

const transactionReturningFunctions = (source: ts.SourceFile): Set<string> => {
  const names = new Set<string>();
  const visit = (node: ts.Node): void => {
    if (
      ts.isFunctionDeclaration(node) &&
      node.name &&
      node.type?.getText(source) === "Transaction"
    ) {
      names.add(node.name.text);
    }
    if (
      ts.isVariableDeclaration(node) &&
      ts.isIdentifier(node.name) &&
      node.initializer &&
      ts.isArrowFunction(node.initializer) &&
      node.initializer.type?.getText(source) === "Transaction"
    )
      names.add(node.name.text);
    ts.forEachChild(node, visit);
  };
  visit(source);
  return names;
};

// Follow a named relative import only when the producer declares Transaction as
// its return type. A DOM method with no arguments is still forbidden.
const importedTransactionFunctions = async (file: string, text: string): Promise<Set<string>> => {
  const source = parseSource(file, text);
  const names = transactionReturningFunctions(source);
  for (const statement of source.statements) {
    if (
      !ts.isImportDeclaration(statement) ||
      !ts.isStringLiteral(statement.moduleSpecifier) ||
      !statement.moduleSpecifier.text.startsWith(".") ||
      !statement.importClause?.namedBindings ||
      !ts.isNamedImports(statement.importClause.namedBindings)
    )
      continue;
    const importedFile = path.resolve(
      ROOT,
      path.dirname(file),
      `${statement.moduleSpecifier.text}.ts`,
    );
    if (!(await Bun.file(importedFile).exists())) continue;
    const importedText = await Bun.file(importedFile).text();
    if (!importedText.includes("Transaction")) continue;
    const exported = transactionReturningFunctions(parseSource(importedFile, importedText));
    for (const binding of statement.importClause.namedBindings.elements) {
      if (exported.has(binding.propertyName?.text ?? binding.name.text))
        names.add(binding.name.text);
    }
  }
  return names;
};

describe("editor scroll-root boundary", () => {
  test("detects DOM calls and every mutation syntax in TS, TSX, and Vue", () => {
    const code = `
      element.scrollIntoView();
      ref.current?.scrollIntoView({ block: "center" });
      viewport.scrollTo({ top: page.offsetTop });
      root["scrollBy"](0, 20);
      viewport.scrollTop = 0;
      root.scrollLeft += 10;
      --root.scrollTop;
      root["scrollLeft"]++;
      const tr = document.querySelector("p");
      tr?.scrollIntoView();
    `;
    for (const file of ["fixture.ts", "fixture.tsx", "fixture.vue"]) {
      const text = file.endsWith(".vue")
        ? `<template><p>scrollIntoView()</p></template>\n<script setup lang="ts">${code}</script>`
        : code;
      expect(scrollViolations({ file, text }).violations).toHaveLength(9);
    }
  });

  test("detects all coordinate assignment operators", () => {
    const operators = [
      "=",
      "+=",
      "-=",
      "*=",
      "/=",
      "%=",
      "**=",
      "<<=",
      ">>=",
      ">>>=",
      "&=",
      "|=",
      "^=",
      "&&=",
      "||=",
      "??=",
    ];
    for (const coordinate of COORDINATES) {
      for (const operator of operators) {
        const text = `root["${coordinate}"] ${operator} 10;`;
        expect(scrollViolations({ file: "fixture.ts", text }).violations).toHaveLength(1);
      }
    }
  });

  test("allows PM transactions and reads but rejects DOM writes in the same scope", () => {
    const text = `
      const tr = view.state.tr;
      view.dispatch(tr.setSelection(selection).scrollIntoView());
      dispatch(state.tr.replaceSelection(slice).scrollIntoView());
      function apply(transaction: Transaction) { return transaction.scrollIntoView(); }
      dispatch(markSectionBreakRemoval(tr, state.doc).scrollIntoView());
      const top = viewport.scrollTop;
      // element.scrollIntoView();
      const prose = "viewport.scrollTop = 0";
      function nested() { const tr = element; tr.scrollIntoView(); }
      element.scrollIntoView();
    `;
    expect(
      scrollViolations({
        file: "fixture.ts",
        text,
        transactionFunctions: new Set(["markSectionBreakRemoval"]),
      }).violations,
    ).toHaveLength(2);
  });

  test("the outline exemption is scoped to the list callback and exact writes", () => {
    const file = "packages/react/src/ui/defaults/outline-rail.tsx";
    const text = `
      useEffect(() => {
        const list = listRef.current;
        list.scrollTop = top;
        list.scrollTop = bottom - list.clientHeight;
        viewport.scrollTop = top;
        list.scrollLeft = 0;
      }, [activeIndex, presentation]);
      list.scrollTop = top;
    `;
    const result = scrollViolations({ file, text });
    expect(result.outlineWrites).toBe(2);
    expect(result.violations).toHaveLength(3);
  });

  test("all production DOM scroll writes use the core scroll-root helper", async () => {
    const violations: string[] = [];
    const outlineWrites: Record<string, number> = {};
    let scanned = 0;
    for await (const file of new Bun.Glob("packages/*/src/**/*.{ts,tsx,vue}").scan({ cwd: ROOT })) {
      if (!isProductionSource(file)) continue;
      const text = await Bun.file(path.join(ROOT, file)).text();
      scanned += 1;
      if (!/scroll(?:IntoView|To|By|Top|Left)/u.test(text)) continue;
      const transactionFunctions = await importedTransactionFunctions(file, text);
      const result = scrollViolations({ file, text, transactionFunctions });
      violations.push(...result.violations);
      if (result.outlineWrites > 0) outlineWrites[file] = result.outlineWrites;
    }
    expect(scanned).toBeGreaterThan(100);
    expect(outlineWrites).toEqual(
      Object.fromEntries(Object.keys(OUTLINE_LISTS).map((file) => [file, 2])),
    );
    expect(violations.sort()).toEqual([]);
  });
});
