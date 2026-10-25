/** DOM scroll writes belong to the shared editor scroll-root boundary. */
import { describe, expect, test } from "bun:test";
import path from "node:path";
import type { TSType } from "oxc-parser";

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
    ? text.replace(
        /<script\b[^>]*>([\s\S]*?)<\/script>|[^<]+|</gu,
        (match, body: unknown) =>
          typeof body === "string"
            ? match.slice(0, match.indexOf(body)).replace(/[^\n]/gu, " ") +
              body +
              match
                .slice(match.indexOf(body) + body.length)
                .replace(/[^\n]/gu, " ")
            : match.replace(/[^\n]/gu, " "),
      )
    : text;
  const { program } = parseLikeTypeScript(
    file,
    script,
    file.endsWith(".tsx") ? "tsx" : "ts",
  );
  return { program, text: script, ...collectSyntaxTree(program) };
};
type Source = ReturnType<typeof parseSource>;

const propertyName = (node: Node): string | undefined => {
  if (node.type !== "MemberExpression") return undefined;
  if (!node.computed)
    return node.property.type === "PrivateIdentifier"
      ? `#${node.property.name}`
      : node.property.name;
  if (
    node.property.type === "Literal" &&
    typeof node.property.value === "string"
  )
    return node.property.value;
  return undefined;
};

const receiver = (node: Node): Node | undefined =>
  node.type === "MemberExpression" ? node.object : undefined;

/** A TypeScript variable, parameter or binding element with an identifier name. */
type Binding = {
  readonly kind: "variable" | "parameter" | "element";
  readonly node: Node;
  readonly name: string;
  readonly type: TSType | undefined;
  readonly initializer: Node | undefined;
  /** An element's written key, as TypeScript's `propertyName.getText()`. */
  readonly key: string | undefined;
};

type ScrollScanOptions = {
  file: string;
  text: string;
  transactionFunctions?: ReadonlySet<string>;
};

const scrollViolations = ({
  file,
  text,
  transactionFunctions = new Set(),
}: ScrollScanOptions) => {
  if (file === SCROLL_WRITER) return { violations: [], outlineWrites: 0 };
  const source = parseSource(file, text);
  const { nodes, parentOf } = source;
  const lineOf = createLineIndex(source.text);
  const scopeOf = (node: Node): Node => {
    let scope = parentOf(node);
    while (
      scope !== undefined &&
      scope.type !== "Program" &&
      !isBlock(scope) &&
      !isFunctionLike(scope)
    ) {
      scope = parentOf(scope);
    }
    return scope ?? source.program;
  };
  const ancestors = (node: Node): Node[] => {
    const result: Node[] = [];
    let current: Node | undefined = node;
    while (current) {
      result.push(current);
      current = parentOf(current);
    }
    return result;
  };

  const bindings: Binding[] = [];
  const bindPattern = (pattern: Node): void => {
    for (const element of bindingElements(pattern)) {
      if (element.name.type === "Identifier")
        bindings.push({
          kind: "element",
          node: element.node,
          name: element.name.name,
          type: undefined,
          initializer: element.initializer,
          key: propertyNameText(source.text, element),
        });
      else bindPattern(element.name);
    }
  };
  const bind = (
    kind: "variable" | "parameter",
    node: Node,
    declared: { name: Node; type: TSType | undefined },
    initializer: Node | undefined,
  ): void => {
    if (declared.name.type === "Identifier")
      bindings.push({
        kind,
        node,
        name: declared.name.name,
        type: declared.type,
        initializer,
        key: undefined,
      });
    else bindPattern(declared.name);
  };
  for (const node of nodes) {
    if (node.type === "VariableDeclarator")
      bind("variable", node, parameterParts(node.id), node.init ?? undefined);
    // TypeScript declares a catch clause's variable like any other.
    if (node.type === "CatchClause" && node.param)
      bind("variable", node.param, parameterParts(node.param), undefined);
    if (isFunctionLike(node))
      for (const parameter of parametersOf(node)) {
        const parts = parameterParts(parameter);
        bind("parameter", parameter, parts, parts.initializer);
      }
  }

  const bindingOf = (node: Node & { name: string }) => {
    const scopes = ancestors(node);
    return bindings
      .filter(
        (candidate) =>
          candidate.name === node.name &&
          candidate.node.start < node.start &&
          scopes.includes(scopeOf(candidate.node)),
      )
      .sort(
        (left, right) =>
          scopes.indexOf(scopeOf(left.node)) -
            scopes.indexOf(scopeOf(right.node)) ||
          right.node.start - left.node.start,
      )
      .at(0);
  };
  const hasTransactionReturn = (type: Node): boolean => {
    let found = false;
    walkAst(type, (node) => {
      if (found) return false;
      if (node.type !== "TSFunctionType") return true;
      found =
        nodeText(source.text, node.returnType.typeAnnotation) === "Transaction";
      return false;
    });
    return found;
  };
  const isTransaction = (node: Node, seen = new Set<Node>()): boolean => {
    if (seen.has(node)) return false;
    seen.add(node);
    if (
      node.type === "ParenthesizedExpression" ||
      node.type === "TSNonNullExpression"
    ) {
      return isTransaction(node.expression, seen);
    }
    if (node.type === "LogicalExpression" && node.operator === "??") {
      return isTransaction(node.left, seen) && isTransaction(node.right, seen);
    }
    if (propertyName(node) === "tr") return true;
    if (node.type === "CallExpression") {
      const object = receiver(node.callee);
      if (object) return isTransaction(object, seen);
      if (node.callee.type !== "Identifier") return false;
      if (transactionFunctions.has(node.callee.name)) return true;
      const binding = bindingOf(node.callee);
      return (
        binding !== undefined &&
        binding.kind !== "element" &&
        binding.type !== undefined &&
        hasTransactionReturn(binding.type)
      );
    }
    if (node.type !== "Identifier") return false;
    const binding = bindingOf(node);
    if (!binding) return false;
    if (binding.kind === "element") {
      return (binding.key ?? binding.name) === "tr";
    }
    if (
      binding.type &&
      /\bTransaction\b/u.test(nodeText(source.text, binding.type))
    )
      return true;
    return (
      binding.initializer !== undefined &&
      isTransaction(binding.initializer, seen)
    );
  };

  const isOutlineListWrite = (
    node: Node & { type: "AssignmentExpression" },
  ): boolean => {
    if (!(file in OUTLINE_LISTS) || propertyName(node.left) !== "scrollTop")
      return false;
    const expectedRef = Object.entries(OUTLINE_LISTS)
      .find(([name]) => name === file)
      ?.at(1);
    const object = receiver(node.left);
    if (!object || nodeText(source.text, object) !== "list") return false;
    if (
      ![
        "list.scrollTop = top",
        "list.scrollTop = bottom - list.clientHeight",
      ].includes(nodeText(source.text, node))
    ) {
      return false;
    }
    const callback = ancestors(node).find(
      (ancestor) => ancestor.type === "ArrowFunctionExpression",
    );
    if (
      !callback ||
      callback.type !== "ArrowFunctionExpression" ||
      callback.body.type !== "BlockStatement"
    )
      return false;
    const call = parentOf(callback);
    if (call?.type !== "CallExpression" || call.callee.type !== "Identifier")
      return false;
    const argumentText = (index: number) => {
      const argument = call.arguments.at(index);
      return argument === undefined
        ? undefined
        : nodeText(source.text, argument);
    };
    const isReact =
      call.callee.name === "useEffect" &&
      argumentText(1) === "[activeIndex, presentation]";
    const isVue =
      call.callee.name === "watch" &&
      argumentText(0) === "[activeIndex, () => props.presentation]";
    if (!isReact && !isVue) return false;
    return callback.body.body.some(
      (statement) =>
        statement.type === "VariableDeclaration" &&
        statement.declarations.some(
          (declaration) =>
            declaration.id.type === "Identifier" &&
            declaration.id.name === "list" &&
            declaration.init !== null &&
            nodeText(source.text, declaration.init) === expectedRef,
        ),
    );
  };

  const violations: string[] = [];
  let outlineWrites = 0;
  const report = (node: Node): void => {
    const { line } = lineOf(node.start);
    violations.push(`${file}:${line + 1}: ${nodeText(source.text, node)}`);
  };
  for (const node of nodes) {
    if (
      node.type === "CallExpression" &&
      METHODS.has(propertyName(node.callee) ?? "")
    ) {
      const object = receiver(node.callee);
      const transactionScroll =
        propertyName(node.callee) === "scrollIntoView" &&
        node.arguments.length === 0 &&
        object !== undefined &&
        isTransaction(object);
      if (!transactionScroll) report(node);
    }
    if (
      node.type === "AssignmentExpression" &&
      COORDINATES.has(propertyName(node.left) ?? "")
    ) {
      if (isOutlineListWrite(node)) outlineWrites += 1;
      else report(node);
    }
    if (
      node.type === "UpdateExpression" &&
      COORDINATES.has(propertyName(node.argument) ?? "")
    )
      report(node);
  }
  return { violations, outlineWrites };
};

const isProductionSource = (file: string): boolean =>
  !/^packages\/playground(?:-vue)?\//u.test(file) &&
  !/(?:^|\/)(?:__tests__|__fixtures__|test|tests)\//u.test(file) &&
  !/\.(?:test|spec|typecheck)\.[^.]+$/u.test(file);

const transactionReturningFunctions = (source: Source): Set<string> => {
  const names = new Set<string>();
  const returnsTransaction = (returnType: { typeAnnotation: Node } | null) =>
    returnType !== null &&
    nodeText(source.text, returnType.typeAnnotation) === "Transaction";
  for (const node of source.nodes) {
    if (
      (node.type === "FunctionDeclaration" ||
        node.type === "TSDeclareFunction") &&
      node.id &&
      returnsTransaction(node.returnType ?? null)
    ) {
      names.add(node.id.name);
    }
    if (
      node.type === "VariableDeclarator" &&
      node.id.type === "Identifier" &&
      node.init &&
      node.init.type === "ArrowFunctionExpression" &&
      returnsTransaction(node.init.returnType ?? null)
    )
      names.add(node.id.name);
  }
  return names;
};

// Follow a named relative import only when the producer declares Transaction as
// its return type. A DOM method with no arguments is still forbidden.
const importedTransactionFunctions = async (
  file: string,
  text: string,
): Promise<Set<string>> => {
  const source = parseSource(file, text);
  const names = transactionReturningFunctions(source);
  for (const statement of source.program.body) {
    if (
      statement.type !== "ImportDeclaration" ||
      !statement.source.value.startsWith(".") ||
      !statement.specifiers.some(
        (specifier) => specifier.type === "ImportSpecifier",
      )
    )
      continue;
    const importedFile = path.resolve(
      ROOT,
      path.dirname(file),
      `${statement.source.value}.ts`,
    );
    if (!(await Bun.file(importedFile).exists())) continue;
    const importedText = await Bun.file(importedFile).text();
    if (!importedText.includes("Transaction")) continue;
    const exported = transactionReturningFunctions(
      parseSource(importedFile, importedText),
    );
    for (const binding of statement.specifiers) {
      if (binding.type !== "ImportSpecifier") continue;
      const imported =
        binding.imported.type === "Literal"
          ? binding.imported.value
          : binding.imported.name;
      if (exported.has(imported)) names.add(binding.local.name);
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
        expect(
          scrollViolations({ file: "fixture.ts", text }).violations,
        ).toHaveLength(1);
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
    for await (const file of new Bun.Glob(
      "packages/*/src/**/*.{ts,tsx,vue}",
    ).scan({ cwd: ROOT })) {
      if (!isProductionSource(file)) continue;
      const text = await Bun.file(path.join(ROOT, file)).text();
      scanned += 1;
      if (!/scroll(?:IntoView|To|By|Top|Left)/u.test(text)) continue;
      const transactionFunctions = await importedTransactionFunctions(
        file,
        text,
      );
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
