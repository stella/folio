import { parse } from "@vue/compiler-sfc";
import { panic } from "better-result";
import { describe, expect, test } from "bun:test";
import { existsSync, readFileSync } from "node:fs";
import path from "node:path";

import {
  createLineIndex,
  declaresBindings,
  type Node,
  nodeText,
  parseSourceModule,
  staticMemberName,
  walkAst,
} from "./lib/oxc-ast";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const HEX_ID_ALLOCATOR = "packages/core/src/utils/hexId.ts";
const NUMERIC_ID_NAME =
  /(?:^id$|id(?:cursor|counter|seed)?$|(?:revision|comment|bookmark|footnote|endnote|part)(?:seed|cursor|counter)$)/i;
const ALLOCATOR_NAME =
  /(?:allocate|generate|next|create).*(?:ids?|revision|comment|bookmark|footnote|endnote|part)$/i;
const STRING_METHODS = new Set([
  "toString",
  "toISOString",
  "toUTCString",
  "toLocaleString",
]);
const OOXML_ID_NAME =
  /(?:revision|comment|bookmark|footnote|endnote|part)(?:id(?:cursor|counter|seed)?|seed|cursor|counter)?$|rsid\w*$|(?:para|text)id(?:cursor|counter|seed)?$/i;

const isNumericIdName = (name: string): boolean =>
  !/undo|relationship|shape|textBox(?:Group|Anchor)|^rId$/i.test(name) &&
  (NUMERIC_ID_NAME.test(name) ||
    ALLOCATOR_NAME.test(name) ||
    OOXML_ID_NAME.test(name));

const stringValue = (node: Node): string | undefined => {
  if (node.type !== "Literal") return undefined;
  const { value } = node;
  return typeof value === "string" ? value : undefined;
};

const propertyName = (node: Node): string | undefined => {
  if (node.type === "Identifier") return node.name;
  return stringValue(node) ?? staticMemberName(node);
};

/** The written name of a property key; a computed key has none. */
const keyName = (node: { computed: boolean; key: Node }) =>
  node.computed || node.key.type === "PrivateIdentifier"
    ? undefined
    : propertyName(node.key);

/** A `.name` member access, as opposed to `[name]`. */
const dotAccess = (node: Node) =>
  node.type === "MemberExpression" && !node.computed ? node : undefined;

const isClockOrRandom = (text: string, node: Node): boolean => {
  const access = dotAccess(node);
  if (access === undefined) return false;
  const name = staticMemberName(access);
  if (name === "getTime") {
    return (
      access.object.type === "NewExpression" &&
      access.object.callee.type === "Identifier" &&
      access.object.callee.name === "Date"
    );
  }
  const owner = nodeText(text, access.object);
  return (
    (name === "now" &&
      [
        "Date",
        "performance",
        "window.performance",
        "globalThis.performance",
      ].includes(owner)) ||
    (name === "random" && owner === "Math")
  );
};

// Functions and signatures: what TypeScript calls function-like.
const FUNCTION_LIKE_NODES = new Set([
  "ArrowFunctionExpression",
  "FunctionDeclaration",
  "FunctionExpression",
  "TSCallSignatureDeclaration",
  "TSConstructSignatureDeclaration",
  "TSConstructorType",
  "TSDeclareFunction",
  "TSEmptyBodyFunctionExpression",
  "TSFunctionType",
  "TSIndexSignature",
  "TSMethodSignature",
]);
// A binding belongs to the nearest block, function or the module.
const SCOPE_NODES = new Set([
  ...FUNCTION_LIKE_NODES,
  "BlockStatement",
  "Program",
  "StaticBlock",
]);
/** Whether any direct child satisfies `predicate`, in source order. */
const someChild = (node: Node, predicate: (child: Node) => boolean) => {
  let found = false;
  walkAst(node, (child) => {
    if (child === node) return true;
    if (!found && predicate(child)) found = true;
    return false;
  });
  return found;
};

type Binding = { declaration: Node; values: Node[] };

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
    result += sourceText
      .slice(cursor, block.loc.start.offset)
      .replace(/[^\r\n]/g, " ");
    result += block.content;
    cursor = block.loc.end.offset;
  }
  return result;
};

/** Follow local aliases to a fixed point, preserving lexical shadowing. */
const forbiddenAllocationSites = (
  file: string,
  sourceText: string,
): string[] => {
  // OOXML hexadecimal rsid/paraId values have a separate, bounded allocator.
  if (file === HEX_ID_ALLOCATOR) return [];
  const text = scriptSource(file, sourceText);
  // Vue scripts are TypeScript-shaped; parse them as TypeScript.
  const { program } = parseSourceModule(
    file.endsWith(".vue") ? `${file}.ts` : file,
    text,
  );
  const lineOf = createLineIndex(text);
  const parents = new Map<Node, Node>();
  walkAst(program, (node, ancestors) => {
    const parent = ancestors.at(-1);
    if (parent !== undefined) parents.set(node, parent);
  });
  const scopes = new Map<Node, Map<string, Binding>>();
  const bindings: Binding[] = [];
  const scopeOf = (node: Node): Node => {
    let scope = parents.get(node);
    while (scope !== undefined && !SCOPE_NODES.has(scope.type)) {
      scope = parents.get(scope);
    }
    return scope ?? program;
  };
  const resolve = (node: Node, name: string): Binding | undefined => {
    let scope: Node | undefined = node;
    while (scope !== undefined) {
      const binding = scopes.get(scope)?.get(name);
      if (binding !== undefined) return binding;
      scope = parents.get(scope);
    }
    return undefined;
  };
  const declare = (declaration: Node, name: string, values: Node[]) => {
    const scope = scopeOf(declaration);
    const names = scopes.get(scope) ?? new Map<string, Binding>();
    const binding = { declaration, values };
    names.set(name, binding);
    scopes.set(scope, names);
    bindings.push(binding);
  };
  walkAst(program, (node) => {
    if (node.type === "VariableDeclarator" && node.id.type === "Identifier") {
      declare(node, node.id.name, node.init === null ? [] : [node.init]);
    } else if (
      (node.type === "FunctionDeclaration" ||
        node.type === "TSDeclareFunction") &&
      node.id !== null
    ) {
      declare(node, node.id.name, [node]);
    } else if (
      node.type === "CatchClause" &&
      node.param?.type === "Identifier"
    ) {
      // A catch variable is a declaration without a value.
      declare(node.param, node.param.name, []);
    }
  });
  walkAst(program, (node) => {
    if (
      node.type === "AssignmentExpression" &&
      node.left.type === "Identifier"
    ) {
      resolve(node.left, node.left.name)?.values.push(node.right);
    }
  });
  const tainted = new Set<Node>();
  // ESTree wraps a method's body in a function value; the method is not a
  // function expression a value can hold.
  const isMethodValue = (node: Node): boolean => {
    const parent = parents.get(node);
    return (
      parent?.type === "MethodDefinition" ||
      parent?.type === "TSAbstractMethodDefinition" ||
      (parent?.type === "Property" && (parent.method || parent.kind !== "init"))
    );
  };
  const returnsSource = (child: Node): boolean => {
    if (FUNCTION_LIKE_NODES.has(child.type)) return false;
    if (child.type === "ReturnStatement") {
      return child.argument !== null && containsSource(child.argument);
    }
    return someChild(child, returnsSource);
  };
  const containsSource = (node: Node): boolean => {
    if (
      node.type === "ArrowFunctionExpression" ||
      node.type === "FunctionDeclaration" ||
      node.type === "TSDeclareFunction" ||
      (node.type === "FunctionExpression" && !isMethodValue(node))
    ) {
      if (!node.body) return false;
      if (node.body.type !== "BlockStatement") return containsSource(node.body);
      return returnsSource(node.body);
    }
    if (isClockOrRandom(text, node)) return true;
    if (node.type === "Identifier") {
      const binding = resolve(node, node.name);
      return binding !== undefined && tainted.has(binding.declaration);
    }
    if (node.type === "Property" && node.kind === "init" && !node.method)
      return containsSource(node.value);
    const access = dotAccess(node);
    if (access !== undefined) return containsSource(access.object);
    return someChild(node, containsSource);
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
  const isStringOutput = (node: Node, visited = new Set<Node>()): boolean => {
    if (visited.has(node)) return false;
    visited.add(node);
    if (node.type === "TemplateLiteral" || stringValue(node) !== undefined)
      return true;
    if (node.type === "Identifier") {
      const values = resolve(node, node.name)?.values;
      return (
        values !== undefined &&
        values.length > 0 &&
        values.every((value) => isStringOutput(value, visited))
      );
    }
    // An optional chain is the call or access it wraps.
    if (
      node.type === "ParenthesizedExpression" ||
      node.type === "TSAsExpression" ||
      node.type === "TSNonNullExpression" ||
      node.type === "ChainExpression"
    )
      return isStringOutput(node.expression, visited);
    if (node.type === "BinaryExpression" && node.operator === "+") {
      return (
        isStringOutput(node.left, visited) ||
        isStringOutput(node.right, visited)
      );
    }
    if (node.type !== "CallExpression") return false;
    if (node.callee.type === "Identifier" && node.callee.name === "String")
      return true;
    const access = dotAccess(node.callee);
    if (access === undefined) return false;
    return (
      STRING_METHODS.has(staticMemberName(access) ?? "") ||
      isStringOutput(access.object, visited)
    );
  };
  const sites: string[] = [];
  const report = (
    node: Node,
    name: string | undefined,
    value: Node | null | undefined,
  ): void => {
    if (
      name === undefined ||
      value === undefined ||
      value === null ||
      !isNumericIdName(name)
    )
      return;
    // Generic application string ids are allowed; OOXML decimal/hexadecimal
    // sinks retain clock taint through string coercion and local aliases.
    if (!OOXML_ID_NAME.test(name) && isStringOutput(value)) return;
    if (!containsSource(value)) return;
    const { line } = lineOf(node.start);
    sites.push(`${file}:${String(line + 1)} (${name})`);
  };
  walkAst(program, (node, ancestors) => {
    if (node.type === "VariableDeclarator") {
      report(
        node,
        node.id.type === "Identifier" ? node.id.name : undefined,
        node.init,
      );
    } else if (
      node.type === "PropertyDefinition" ||
      node.type === "TSAbstractPropertyDefinition" ||
      node.type === "AccessorProperty" ||
      node.type === "TSAbstractAccessorProperty"
    ) {
      report(node, keyName(node), node.value);
    } else if (
      node.type === "Property" &&
      // Destructuring that declares bindings has no property assignments.
      (ancestors.at(-1)?.type === "ObjectExpression" ||
        !declaresBindings(ancestors))
    ) {
      if (node.shorthand) {
        report(
          node,
          keyName(node),
          node.value.type === "AssignmentPattern"
            ? node.value.left
            : node.value,
        );
      } else if (node.kind === "init" && !node.method) {
        report(node, keyName(node), node.value);
      }
    } else if (node.type === "AssignmentExpression") {
      report(node, propertyName(node.left), node.right);
    } else if (
      node.type === "FunctionDeclaration" ||
      node.type === "TSDeclareFunction"
    ) {
      report(node, node.id?.name, node);
    }
  });
  return sites;
};

const sourceFiles = (): string[] =>
  [
    ...new Bun.Glob("*/src/**/*.{ts,tsx,vue,js,jsx}").scanSync({
      cwd: path.join(REPO_ROOT, "packages"),
      dot: true,
      onlyFiles: true,
    }),
  ]
    .map((file) => `packages/${file.replaceAll("\\", "/")}`)
    .filter((file) => !/(?:^|\/)node_modules\//.test(file))
    .filter(
      (file) =>
        /^packages\/[^/]+\/src\//.test(file) && !/\.test\.tsx?$/.test(file),
    )
    .toSorted();

const readSource = (file: string): string | undefined =>
  existsSync(file)
    ? // A byte-order mark is not source text.
      readFileSync(file, "utf8").replace(/^\uFEFF/u, "")
    : undefined;

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
    expect(forbiddenAllocationSites("probe.ts", source).length).toBeGreaterThan(
      0,
    );
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
    expect(files).toContain(
      "packages/vue/src/composables/useDocxEditorRefApi.ts",
    );
    expect(files).toContain("packages/vue/src/components/DocxEditor.vue");
    expect(files.some((file) => file.includes("/node_modules/"))).toBe(false);
    expect(
      forbiddenAllocationSites(HEX_ID_ALLOCATOR, "const id = Math.random();"),
    ).toEqual([]);
    const forbidden = files.flatMap((file) => {
      const source = readSource(path.join(REPO_ROOT, file));
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
