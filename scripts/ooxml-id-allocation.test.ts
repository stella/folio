import { describe, expect, test } from "bun:test";
import { parse } from "@vue/compiler-sfc";
import { panic } from "better-result";
import path from "node:path";
import ts from "typescript";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const HEX_ID_ALLOCATOR = "packages/core/src/utils/hexId.ts";
const NUMERIC_ID_NAME =
  /(?:^id$|id(?:cursor|counter|seed)?$|(?:revision|comment|bookmark|footnote|endnote|part)(?:seed|cursor|counter)$)/i;
const ALLOCATOR_NAME =
  /(?:allocate|generate|next|create).*(?:ids?|revision|comment|bookmark|footnote|endnote|part)$/i;
const STRING_METHODS = new Set(["toString", "toISOString", "toUTCString", "toLocaleString"]);
const OOXML_ID_NAME =
  /(?:revision|comment|bookmark|footnote|endnote|part)(?:id(?:cursor|counter|seed)?|seed|cursor|counter)?$|rsid\w*$|(?:para|text)id(?:cursor|counter|seed)?$/i;

const isNumericIdName = (name: string): boolean =>
  !/undo|relationship|shape|textBox(?:Group|Anchor)|^rId$/i.test(name) &&
  (NUMERIC_ID_NAME.test(name) || ALLOCATOR_NAME.test(name) || OOXML_ID_NAME.test(name));

const propertyName = (node: ts.Node): string | undefined => {
  if (ts.isIdentifier(node) || ts.isStringLiteral(node)) return node.text;
  if (ts.isPropertyAccessExpression(node)) return node.name.text;
  if (ts.isElementAccessExpression(node) && ts.isStringLiteral(node.argumentExpression)) {
    return node.argumentExpression.text;
  }
  return undefined;
};

const isClockOrRandom = (node: ts.Node): boolean => {
  if (!ts.isPropertyAccessExpression(node)) return false;
  if (node.name.text === "getTime") {
    return (
      ts.isNewExpression(node.expression) &&
      ts.isIdentifier(node.expression.expression) &&
      node.expression.expression.text === "Date"
    );
  }
  const owner = node.expression.getText();
  return (
    (node.name.text === "now" &&
      ["Date", "performance", "window.performance", "globalThis.performance"].includes(owner)) ||
    (node.name.text === "random" && owner === "Math")
  );
};

type Binding = { declaration: ts.Node; values: ts.Node[] };

const scriptSource = (file: string, sourceText: string): string => {
  if (!file.endsWith(".vue")) return sourceText;
  const { descriptor, errors } = parse(sourceText, { filename: file });
  if (errors.length > 0) panic(`Cannot parse ${file}.`);
  const blocks = [descriptor.script, descriptor.scriptSetup]
    .filter((block) => block !== null)
    .toSorted((left, right) => left.loc.start.offset - right.loc.start.offset);
  let cursor = 0;
  let result = "";
  for (const block of blocks) {
    result += sourceText.slice(cursor, block.loc.start.offset).replace(/[^\r\n]/g, " ");
    result += block.content;
    cursor = block.loc.end.offset;
  }
  return result;
};

/** Follow local aliases to a fixed point, preserving lexical shadowing. */
const forbiddenAllocationSites = (file: string, sourceText: string): string[] => {
  // OOXML hexadecimal rsid/paraId values have a separate, bounded allocator.
  if (file === HEX_ID_ALLOCATOR) return [];
  const source = ts.createSourceFile(
    file,
    scriptSource(file, sourceText),
    ts.ScriptTarget.Latest,
    true,
  );
  const scopes = new Map<ts.Node, Map<string, Binding>>();
  const bindings: Binding[] = [];
  const scopeOf = (node: ts.Node): ts.Node => {
    let scope = node.parent;
    while (!ts.isSourceFile(scope) && !ts.isBlock(scope) && !ts.isFunctionLike(scope)) {
      scope = scope.parent;
    }
    return scope;
  };
  const resolve = (node: ts.Identifier): Binding | undefined => {
    let scope: ts.Node | undefined = node;
    while (scope !== undefined) {
      const binding = scopes.get(scope)?.get(node.text);
      if (binding !== undefined) return binding;
      scope = scope.parent;
    }
    return undefined;
  };
  const collect = (node: ts.Node): void => {
    if (
      (ts.isVariableDeclaration(node) || ts.isFunctionDeclaration(node)) &&
      node.name !== undefined &&
      ts.isIdentifier(node.name)
    ) {
      const scope = scopeOf(node);
      const names = scopes.get(scope) ?? new Map<string, Binding>();
      const value = ts.isVariableDeclaration(node) ? node.initializer : node;
      const binding = { declaration: node, values: value === undefined ? [] : [value] };
      names.set(node.name.text, binding);
      scopes.set(scope, names);
      bindings.push(binding);
    }
    ts.forEachChild(node, collect);
  };
  collect(source);
  const collectAssignments = (node: ts.Node): void => {
    if (
      ts.isBinaryExpression(node) &&
      node.operatorToken.kind >= ts.SyntaxKind.FirstAssignment &&
      node.operatorToken.kind <= ts.SyntaxKind.LastAssignment &&
      ts.isIdentifier(node.left)
    ) {
      resolve(node.left)?.values.push(node.right);
    }
    ts.forEachChild(node, collectAssignments);
  };
  collectAssignments(source);
  const tainted = new Set<ts.Node>();
  const containsSource = (node: ts.Node): boolean => {
    if (
      ts.isArrowFunction(node) ||
      ts.isFunctionExpression(node) ||
      ts.isFunctionDeclaration(node)
    ) {
      if (node.body === undefined) return false;
      if (!ts.isBlock(node.body)) return containsSource(node.body);
      const returnsSource = (child: ts.Node): boolean => {
        if (ts.isFunctionLike(child)) return false;
        if (ts.isReturnStatement(child)) {
          return child.expression !== undefined && containsSource(child.expression);
        }
        return ts.forEachChild(child, (nested) => returnsSource(nested) || undefined) === true;
      };
      return returnsSource(node.body);
    }
    if (isClockOrRandom(node)) return true;
    if (ts.isIdentifier(node)) {
      const binding = resolve(node);
      return binding !== undefined && tainted.has(binding.declaration);
    }
    if (ts.isPropertyAssignment(node)) return containsSource(node.initializer);
    if (ts.isPropertyAccessExpression(node)) return containsSource(node.expression);
    return ts.forEachChild(node, (child) => containsSource(child) || undefined) === true;
  };
  let changed = true;
  while (changed) {
    changed = false;
    for (const { declaration, values } of bindings) {
      if (tainted.has(declaration) || !values.some(containsSource)) continue;
      tainted.add(declaration);
      changed = true;
    }
  }
  const isStringOutput = (node: ts.Node, visited = new Set<ts.Node>()): boolean => {
    if (visited.has(node)) return false;
    visited.add(node);
    if (
      ts.isTemplateExpression(node) ||
      ts.isNoSubstitutionTemplateLiteral(node) ||
      ts.isStringLiteral(node)
    )
      return true;
    if (ts.isIdentifier(node)) {
      const values = resolve(node)?.values;
      return (
        values !== undefined &&
        values.length > 0 &&
        values.every((value) => isStringOutput(value, visited))
      );
    }
    if (
      ts.isParenthesizedExpression(node) ||
      ts.isAsExpression(node) ||
      ts.isNonNullExpression(node)
    )
      return isStringOutput(node.expression, visited);
    if (ts.isBinaryExpression(node) && node.operatorToken.kind === ts.SyntaxKind.PlusToken) {
      return isStringOutput(node.left, visited) || isStringOutput(node.right, visited);
    }
    if (!ts.isCallExpression(node)) return false;
    if (ts.isIdentifier(node.expression) && node.expression.text === "String") return true;
    if (!ts.isPropertyAccessExpression(node.expression)) return false;
    return (
      STRING_METHODS.has(node.expression.name.text) ||
      isStringOutput(node.expression.expression, visited)
    );
  };
  const sites: string[] = [];
  const report = (node: ts.Node, name: string | undefined, value: ts.Node | undefined): void => {
    if (name === undefined || value === undefined || !isNumericIdName(name)) return;
    // Generic application string ids are allowed; OOXML decimal/hexadecimal
    // sinks retain clock taint through string coercion and local aliases.
    if (!OOXML_ID_NAME.test(name) && isStringOutput(value)) return;
    if (!containsSource(value)) return;
    const { line } = source.getLineAndCharacterOfPosition(node.getStart(source));
    sites.push(`${file}:${String(line + 1)} (${name})`);
  };
  const visit = (node: ts.Node): void => {
    if (ts.isVariableDeclaration(node) || ts.isPropertyDeclaration(node)) {
      report(node, propertyName(node.name), node.initializer);
    } else if (ts.isPropertyAssignment(node)) {
      report(node, propertyName(node.name), node.initializer);
    } else if (ts.isShorthandPropertyAssignment(node)) {
      report(node, node.name.text, node.name);
    } else if (
      ts.isBinaryExpression(node) &&
      node.operatorToken.kind >= ts.SyntaxKind.FirstAssignment &&
      node.operatorToken.kind <= ts.SyntaxKind.LastAssignment
    ) {
      report(node, propertyName(node.left), node.right);
    } else if (ts.isFunctionDeclaration(node)) {
      report(node, node.name?.text, node);
    }
    ts.forEachChild(node, visit);
  };
  visit(source);
  return sites;
};

const sourceFiles = (): string[] =>
  ts.sys
    .readDirectory(
      path.join(REPO_ROOT, "packages"),
      [".ts", ".tsx", ".vue", ".js", ".jsx"],
      ["**/node_modules/**"],
    )
    .map((file) => path.relative(REPO_ROOT, file).replaceAll("\\", "/"))
    .filter((file) => /^packages\/[^/]+\/src\//.test(file) && !/\.test\.tsx?$/.test(file))
    .toSorted();

describe("numeric OOXML id allocation", () => {
  test.each([
    "let commentIdCursor = Date.now();",
    "let revisionIdCursor = Date.now() * 1000;",
    "const seed = performance.now(); const commentId = seed + 1;",
    "const seed = Math.floor(Math.random() * 2147483647); const bookmarkId = seed;",
    "const clock = Date.now; const seed = clock(); const footnoteId = seed;",
    "const seed = () => window.performance.now(); const partId = seed();",
    "let seed = 0; seed = Date.now(); const endnoteId = seed;",
    "const seed = Date.now(); const alias = seed; const entry = { id: alias + 1 };",
    "const now = globalThis.performance.now(); target['revisionId'] = now;",
    "function nextCommentId() { return Date.now(); }",
    "const seed = new Date().getTime(); const commentId = seed;",
    "const commentId = String(Date.now());",
    "const id = Number(String(Date.now()));",
    "const seed = String(Date.now()); const revisionId = Number(seed);",
    "const rsid = Date.now().toString(16);",
    "const seed = performance.now().toString(16); const paraId = seed.toUpperCase();",
    "const clock = Date.now; const seed = String(clock()); const textId = seed;",
  ])("rejects numeric clock/random allocation: %s", (source) => {
    expect(forbiddenAllocationSites("probe.ts", source).length).toBeGreaterThan(0);
  });

  test.each([
    "const startedAt = performance.now(); recordTiming(performance.now() - startedAt);",
    "let undoHandleCursor = Date.now(); const documentOperationUndoHandleId = undoHandleCursor++;",
    "const rId = `rId_img_${Date.now()}_${Math.random()}`;",
    "const relationshipId = Date.now().toString();",
    "const bookmarkName = `_Toc${Math.floor(Math.random() * 900000000)}`;",
    "const shapeId = Math.random().toString(36).slice(2);",
    "const seed = Date.now(); function allocate() { const seed = 1; const commentId = seed++; }",
    "let nextCommentId = 1; const comment = { id: nextCommentId++ };",
    "const id = `error-${Date.now()}`;",
    "const seed = Math.random().toString(36); const id = `anchor:${seed}`;",
    "const nextTextBoxGroupId = () => `${Math.random().toString(36)}:0`; const textBoxGroupId = nextTextBoxGroupId();",
    "function createHiddenEditorState() { const started = performance.now(); recordTiming(started); }",
    "function nextCommentId() { const started = Date.now(); recordTiming(started); return 1; }",
    "const nextCommentId = () => { const started = Date.now(); recordTiming(started); return 1; };",
  ])("permits timing, strings, handles, and bounded counters: %s", (source) => {
    expect(forbiddenAllocationSites("probe.ts", source)).toEqual([]);
  });

  test("scans every package src without workspace symlinks and rejects forbidden producers", () => {
    const files = sourceFiles();
    expect(files).toContain("packages/core/src/ai-edits/headless.ts");
    expect(files).toContain("packages/react/src/components/commentsHelpers.ts");
    expect(files).toContain("packages/vue/src/composables/useDocxEditorRefApi.ts");
    expect(files).toContain("packages/vue/src/components/DocxEditor.vue");
    expect(files.some((file) => file.includes("/node_modules/"))).toBe(false);
    expect(forbiddenAllocationSites(HEX_ID_ALLOCATOR, "const id = Math.random();")).toEqual([]);
    const forbidden = files.flatMap((file) => {
      const source = ts.sys.readFile(path.join(REPO_ROOT, file));
      if (source === undefined) panic(`Cannot read ${file}.`);
      return forbiddenAllocationSites(file, source);
    });
    expect(forbidden).toEqual([]);
  }, 30_000);

  test("reads Vue script blocks while preserving source locations", () => {
    expect(
      forbiddenAllocationSites(
        "probe.vue",
        '<template><div>{{ id }}</div></template>\n<script setup lang="ts">\nconst commentId = Date.now();\n</script>',
      ),
    ).toEqual(["probe.vue:3 (commentId)"]);
  });
});
