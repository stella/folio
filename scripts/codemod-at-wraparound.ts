#!/usr/bin/env bun
// Rewrite `.at()` reads of the item before an index in package source to the
// package's `itemBefore` helper, which `folio-at-wraparound` requires.
//
//   xs.at(i - 1)                       -> itemBefore(xs, i)
//   i > 0 ? xs.at(i - 1) : undefined   -> itemBefore(xs, i)
//   i === 0 ? undefined : xs.at(i - 1) -> itemBefore(xs, i)
//
// `.at()` counts a negative index from the end, so `xs.at(i - 1)` at `i === 0`
// reads the last item. `itemBefore` reads the same item whenever `i - 1` is
// not negative and `undefined` when it is, so a guard that only kept the index
// from going negative folds into it. Any other flagged shape (another
// subtrahend, an optional chain, a computed negation) stops the run for a hand
// rewrite; a read kept on purpose carries a disable comment and is skipped.
// Running it again changes nothing.
//
// Usage: bun scripts/codemod-at-wraparound.ts [--check]

import { readFileSync, writeFileSync } from "node:fs";
import path from "node:path";

import { wrappingIndexOf } from "../.oxlint-plugins/folio-at-wraparound";
import {
  createLineIndex,
  type Node,
  nodeText,
  parseSourceModule,
  walkAst,
} from "./lib/oxc-ast";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const RULE = "folio-at-wraparound/no-computed-negative-index";

// One helper module per package; a package with a flagged read and no entry
// here stops the run until it gets one.
const HELPERS: Readonly<Record<string, string>> = {
  core: "packages/core/src/utils/listItems.ts",
  "docx-core": "packages/docx-core/src/ops/listItems.ts",
  vue: "packages/vue/src/utils/listItems.ts",
};

const HELPER = "itemBefore";

type Rewrite = {
  readonly start: number;
  readonly end: number;
  readonly replacement: string;
};

const isSourceFile = (relative: string): boolean =>
  !/\.(?:test|spec)\.[cm]?tsx?$/u.test(relative) &&
  !relative.endsWith(".d.ts") &&
  !relative.endsWith(".gen.ts") &&
  !relative.split("/").includes("__tests__") &&
  !relative.split("/").includes("generated");

const unwrap = (node: Node): Node =>
  node.type === "ParenthesizedExpression" || node.type === "TSNonNullExpression"
    ? unwrap(node.expression)
    : node;

const isNumber = (node: Node, value: number): boolean =>
  node.type === "Literal" && node.value === value;

const isUndefined = (node: Node): boolean =>
  node.type === "Identifier" && node.name === "undefined";

// Whether an optional link earlier in the same chain can skip the read:
// `a?.list.at(i - 1)` reads nothing when `a` is nullish, which a helper call
// over `a?.list` would not.
const continuesOptionalChain = (node: Node): boolean => {
  switch (node.type) {
    case "MemberExpression":
      return node.optional || continuesOptionalChain(node.object);
    case "CallExpression":
      return node.optional || continuesOptionalChain(node.callee);
    case "TSNonNullExpression":
      return continuesOptionalChain(node.expression);
    default:
      return false;
  }
};

type FlaggedRead = {
  readonly call: Node;
  // The `at` name, whose line a disable comment covers.
  readonly property: Node;
  readonly index: Node;
};

// A `.at()` call the rule flags, as a subtraction the codemod rewrites, or
// nothing. The selection is the rule's own.
const flaggedRead = (relative: string, node: Node): FlaggedRead | undefined => {
  const kind = wrappingIndexOf(node);
  if (kind === undefined) return undefined;
  if (node.type !== "CallExpression" || node.callee.type !== "MemberExpression")
    throw new Error(`${relative}:${node.start}: not a member call`);
  const [argument] = node.arguments;
  if (
    kind === "negation" ||
    argument === undefined ||
    node.optional ||
    node.callee.computed ||
    continuesOptionalChain(node.callee)
  ) {
    throw new Error(`${relative}:${node.start}: rewrite this read by hand`);
  }
  return {
    call: node,
    property: node.callee.property,
    index: unwrap(argument),
  };
};

type Guarded = { readonly start: number; readonly end: number };

// `i > 0 ? call : undefined` or `i === 0 ? undefined : call`, where `call`
// reads `xs.at(i - 1)`: the span of the whole conditional.
const guardAround = (
  text: string,
  call: Node,
  subtrahend: Node,
  parent: Node | undefined,
): Guarded | undefined => {
  if (parent?.type !== "ConditionalExpression") return undefined;
  const test = unwrap(parent.test);
  if (test.type !== "BinaryExpression" || !isNumber(unwrap(test.right), 0)) {
    return undefined;
  }
  if (
    nodeText(text, unwrap(test.left)) !== nodeText(text, unwrap(subtrahend))
  ) {
    return undefined;
  }
  const consequent = unwrap(parent.consequent);
  const alternate = unwrap(parent.alternate);
  const guards =
    (test.operator === ">" && consequent === call && isUndefined(alternate)) ||
    (test.operator === "===" && isUndefined(consequent) && alternate === call);
  return guards ? { start: parent.start, end: parent.end } : undefined;
};

type RewriteCallOptions = {
  readonly relative: string;
  readonly text: string;
  readonly call: Node;
  readonly index: Node;
  readonly parent: Node | undefined;
};

const rewriteCall = ({
  relative,
  text,
  call,
  index,
  parent,
}: RewriteCallOptions): Rewrite => {
  if (call.type !== "CallExpression" || call.callee.type !== "MemberExpression")
    throw new Error("not a member call");
  if (index.type !== "BinaryExpression") throw new Error("not a subtraction");
  if (!isNumber(unwrap(index.right), 1)) {
    throw new Error(
      `${relative}:${call.start}: only \`- 1\` has a helper; rewrite this read by hand`,
    );
  }
  const list = nodeText(text, call.callee.object);
  // `(a ? b : c) - 1` needs its parentheses; as an argument it does not.
  const { left } = index;
  const position = nodeText(
    text,
    left.type === "ParenthesizedExpression" &&
      left.expression.type !== "SequenceExpression"
      ? left.expression
      : left,
  );
  const span = guardAround(text, call, left, parent) ?? call;
  return {
    start: span.start,
    end: span.end,
    replacement: `${HELPER}(${list}, ${position})`,
  };
};

// Flagged reads that contain no other flagged read; an enclosing one is
// rewritten on the next pass, over the rewritten text.
const innermostRewrites = (relative: string, text: string): Rewrite[] => {
  const { program, comments } = parseSourceModule(relative, text);
  const line = createLineIndex(text);
  // Lines under a disable comment for the rule hold reads kept on purpose.
  const kept = new Set(
    comments
      .filter(
        ({ value }) =>
          value.includes("oxlint-disable-next-line") && value.includes(RULE),
      )
      .map(({ end }) => line(end).line + 1),
  );
  const calls: { node: Node; index: Node; parent: Node | undefined }[] = [];
  walkAst(program, (node, ancestors) => {
    if (
      node.type === "CallExpression" &&
      node.callee.type === "MemberExpression" &&
      kept.has(line(node.callee.property.start).line)
    )
      return;
    const read = flaggedRead(relative, node);
    if (read === undefined) return;
    const { index } = read;
    const parent = ancestors.findLast(
      (ancestor) =>
        ancestor.type !== "ParenthesizedExpression" &&
        ancestor.type !== "TSNonNullExpression",
    );
    calls.push({ node, index, parent });
  });
  return calls
    .filter(
      ({ node }) =>
        !calls.some(
          ({ node: other }) =>
            other !== node &&
            other.start >= node.start &&
            other.end <= node.end,
        ),
    )
    .map(({ node, index, parent }) =>
      rewriteCall({ relative, text, call: node, index, parent }),
    );
};

const importSpecifier = (file: string, helperFile: string): string => {
  const relative = path
    .relative(path.dirname(file), helperFile)
    .replace(/\.tsx?$/u, "");
  return relative.startsWith(".") ? relative : `./${relative}`;
};

// Imports sort by specifier: a new declaration goes before the first relative
// import that sorts after it.
const addImport = (relative: string, text: string, specifier: string) => {
  const { program } = parseSourceModule(relative, text);
  const imports = program.body.filter(
    (statement) => statement.type === "ImportDeclaration",
  );
  const existing = imports.find(
    (statement) =>
      statement.source.value === specifier && statement.importKind !== "type",
  );
  const imported = new Set(
    (existing?.specifiers ?? []).map(
      (specifierNode) => specifierNode.local.name,
    ),
  );
  if (imported.has(HELPER)) return text;
  const lastSpecifier = existing?.specifiers.at(-1);
  if (lastSpecifier !== undefined) {
    const at = lastSpecifier.end;
    return `${text.slice(0, at)}, ${HELPER}${text.slice(at)}`;
  }
  const statement = `import { ${HELPER} } from "${specifier}";\n`;
  const next = imports.find(
    (declaration) =>
      declaration.source.value.startsWith(".") &&
      declaration.source.value.localeCompare(specifier) > 0,
  );
  if (next !== undefined) {
    return `${text.slice(0, next.start)}${statement}${text.slice(next.start)}`;
  }
  const last = imports.at(-1);
  if (last === undefined) return `${statement}${text}`;
  const lineEnd = text.indexOf("\n", last.end) + 1;
  return `${text.slice(0, lineEnd)}${statement}${text.slice(lineEnd)}`;
};

type FileResult = {
  readonly text: string;
  readonly sites: readonly string[];
};

const rewriteFile = (relative: string, original: string): FileResult => {
  const absolute = path.join(REPO_ROOT, relative);
  const packageName = relative.split("/")[1] ?? "";
  const sites: string[] = [];
  let text = original;
  for (;;) {
    const rewrites = innermostRewrites(relative, text);
    if (rewrites.length === 0) break;
    const line = createLineIndex(text);
    for (const rewrite of rewrites.toSorted((a, b) => b.start - a.start)) {
      sites.push(
        `${relative}:${line(rewrite.start).line + 1}: ${text.slice(rewrite.start, rewrite.end)} -> ${rewrite.replacement}`,
      );
      text = `${text.slice(0, rewrite.start)}${rewrite.replacement}${text.slice(rewrite.end)}`;
    }
  }
  if (sites.length === 0) return { text, sites };
  const helperFile = HELPERS[packageName];
  if (helperFile === undefined) {
    throw new Error(
      `${relative}: packages/${packageName} has no ${HELPER} helper; add one to HELPERS`,
    );
  }
  const helperSource = readFileSync(path.join(REPO_ROOT, helperFile), "utf8");
  const specifier = importSpecifier(absolute, path.join(REPO_ROOT, helperFile));
  if (!helperSource.includes(`export const ${HELPER} =`)) {
    throw new Error(`${helperFile} does not export ${HELPER}`);
  }
  // A name the file already uses for something else would capture the call.
  if (
    new RegExp(`\\b${HELPER}\\b`, "u").test(original) &&
    !original.includes(`from "${specifier}"`)
  ) {
    throw new Error(`${relative}: \`${HELPER}\` is already bound`);
  }
  return { text: addImport(relative, text, specifier), sites };
};

const check = process.argv.includes("--check");
const changed: string[] = [];
const helperFiles = new Set(Object.values(HELPERS));
for (const relative of new Bun.Glob("packages/*/src/**/*.{ts,tsx}").scanSync({
  cwd: REPO_ROOT,
})) {
  if (!isSourceFile(relative) || helperFiles.has(relative)) continue;
  const original = readFileSync(path.join(REPO_ROOT, relative), "utf8");
  if (!original.includes(".at(")) continue;
  const { text, sites } = rewriteFile(relative, original);
  if (text === original) continue;
  changed.push(relative);
  for (const site of sites.toReversed()) console.log(site);
  if (!check) writeFileSync(path.join(REPO_ROOT, relative), text);
}

console.log(
  `${changed.length} file(s) ${check ? "would change" : "rewritten"}`,
);
if (check && changed.length > 0) process.exit(1);
