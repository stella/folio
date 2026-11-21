#!/usr/bin/env bun
// Rewrite spread calls into `Math.max` / `Math.min` in package source to the
// package's `maxOf` / `minOf` helper, which `folio-math-extremes` requires.
//
//   Math.max(...xs)          -> maxOf(xs)
//   Math.max(0, ...xs)       -> maxOf(xs, 0)
//   Math.min(...xs, limit)   -> minOf(xs, limit)
//   Math.max(a, ...xs, ...ys) -> maxOf([a, ...xs, ...ys])
//
// `maxOf(values, floor)` is `Math.max(floor, ...values)`, so each rewrite keeps
// the call's result, the empty-input case included. A lone fixed argument
// moves after the values only when reading it has no effect (a literal, a
// name, a property read); any other list keeps its order inside an array.
// Running it again changes nothing.
//
// Usage: bun scripts/codemod-math-extremes.ts [--check]

import { readFileSync, writeFileSync } from "node:fs";
import path from "node:path";

import {
  createLineIndex,
  type Node,
  nodeText,
  parseSourceModule,
  walkAst,
} from "./lib/oxc-ast";

const REPO_ROOT = path.resolve(import.meta.dir, "..");

// One helper module per package; a package with a spread call and no entry
// here stops the run until it gets one.
const HELPERS: Readonly<Record<string, string>> = {
  cli: "packages/cli/src/extremes.ts",
  core: "packages/core/src/utils/extremes.ts",
  "docx-core": "packages/docx-core/src/model/extremes.ts",
};

const HELPER_NAMES = { max: "maxOf", min: "minOf" } as const;

type Extreme = keyof typeof HELPER_NAMES;

type Rewrite = {
  readonly start: number;
  readonly end: number;
  readonly replacement: string;
  readonly helper: (typeof HELPER_NAMES)[Extreme];
};

const isSourceFile = (relative: string): boolean =>
  !/\.(?:test|spec)\.[cm]?tsx?$/u.test(relative) &&
  !relative.endsWith(".d.ts") &&
  !relative.endsWith(".gen.ts") &&
  !relative.split("/").includes("__tests__") &&
  !relative.split("/").includes("generated");

const extremeOf = (node: Node): Extreme | undefined => {
  if (node.type !== "CallExpression") return undefined;
  const { callee } = node;
  if (callee.type !== "MemberExpression" || callee.computed) return undefined;
  if (callee.object.type !== "Identifier" || callee.object.name !== "Math") {
    return undefined;
  }
  const { name } = callee.property;
  if (name !== "max" && name !== "min") return undefined;
  return node.arguments.some((arg) => arg.type === "SpreadElement")
    ? name
    : undefined;
};

// Reading the expression runs no code of the file's own: moving it after the
// values cannot change what either produces.
const isInert = (node: Node): boolean => {
  switch (node.type) {
    case "Literal":
    case "Identifier":
      return true;
    case "ParenthesizedExpression":
    case "TSNonNullExpression":
    case "ChainExpression":
      return isInert(node.expression);
    case "UnaryExpression":
      return node.operator === "-" && isInert(node.argument);
    case "LogicalExpression":
      return isInert(node.left) && isInert(node.right);
    case "MemberExpression":
      return (
        isInert(node.object) &&
        (!node.computed || node.property.type === "Literal")
      );
    default:
      return false;
  }
};

const rewriteCall = (text: string, call: Node, extreme: Extreme): Rewrite => {
  if (call.type !== "CallExpression") throw new Error("not a call");
  const helper = HELPER_NAMES[extreme];
  const args = call.arguments;
  const spreads = args.filter((arg) => arg.type === "SpreadElement");
  const fixed = args.filter((arg) => arg.type !== "SpreadElement");
  const [spread] = spreads;
  const [bound] = fixed;
  let replacement: string;
  if (
    spreads.length === 1 &&
    spread !== undefined &&
    spread.type === "SpreadElement" &&
    fixed.length <= 1 &&
    fixed.every(isInert)
  ) {
    // `...(xs ?? [])` needs its parentheses; as an argument it does not.
    const values = nodeText(
      text,
      spread.argument.type === "ParenthesizedExpression" &&
        spread.argument.expression.type !== "SequenceExpression"
        ? spread.argument.expression
        : spread.argument,
    );
    replacement =
      bound === undefined
        ? `${helper}(${values})`
        : `${helper}(${values}, ${nodeText(text, bound)})`;
  } else {
    const list = args.map((arg) => nodeText(text, arg)).join(", ");
    replacement = `${helper}([${list}])`;
  }
  return { start: call.start, end: call.end, replacement, helper };
};

// Spread calls that contain no other spread call; an enclosing one is
// rewritten on the next pass, over the rewritten text.
const innermostRewrites = (fileName: string, text: string): Rewrite[] => {
  const { program } = parseSourceModule(fileName, text);
  const calls: { node: Node; extreme: Extreme }[] = [];
  walkAst(program, (node) => {
    const extreme = extremeOf(node);
    if (extreme !== undefined) calls.push({ node, extreme });
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
    .map(({ node, extreme }) => rewriteCall(text, node, extreme));
};

const importSpecifier = (file: string, helperFile: string): string => {
  const relative = path
    .relative(path.dirname(file), helperFile)
    .replace(/\.ts$/u, "");
  return relative.startsWith(".") ? relative : `./${relative}`;
};

const addImports = (
  fileName: string,
  text: string,
  specifier: string,
  names: ReadonlySet<string>,
): string => {
  const { program } = parseSourceModule(fileName, text);
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
  const missing = [...names].filter((name) => !imported.has(name)).sort();
  if (missing.length === 0) return text;
  const lastSpecifier = existing?.specifiers.at(-1);
  if (lastSpecifier !== undefined) {
    const at = lastSpecifier.end;
    return `${text.slice(0, at)}, ${missing.join(", ")}${text.slice(at)}`;
  }
  const statement = `import { ${missing.join(", ")} } from "${specifier}";\n`;
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
  const helpers = new Set<string>();
  let text = original;
  for (;;) {
    const rewrites = innermostRewrites(relative, text);
    if (rewrites.length === 0) break;
    const line = createLineIndex(text);
    for (const rewrite of rewrites.toSorted((a, b) => b.start - a.start)) {
      sites.push(
        `${relative}:${line(rewrite.start).line + 1}: ${text.slice(rewrite.start, rewrite.end)} -> ${rewrite.replacement}`,
      );
      helpers.add(rewrite.helper);
      text = `${text.slice(0, rewrite.start)}${rewrite.replacement}${text.slice(rewrite.end)}`;
    }
  }
  if (helpers.size === 0) return { text, sites };
  const helperFile = HELPERS[packageName];
  if (helperFile === undefined) {
    throw new Error(
      `${relative}: packages/${packageName} has no maxOf/minOf helper; add one to HELPERS`,
    );
  }
  const helperSource = readFileSync(path.join(REPO_ROOT, helperFile), "utf8");
  const specifier = importSpecifier(absolute, path.join(REPO_ROOT, helperFile));
  for (const helper of helpers) {
    if (!helperSource.includes(`export const ${helper} =`)) {
      throw new Error(`${helperFile} does not export ${helper}`);
    }
    // A name the file already uses for something else would capture the call.
    if (
      new RegExp(`\\b${helper}\\b`, "u").test(original) &&
      !original.includes(`from "${specifier}"`)
    ) {
      throw new Error(`${relative}: \`${helper}\` is already bound`);
    }
  }
  return { text: addImports(relative, text, specifier, helpers), sites };
};

const check = process.argv.includes("--check");
const changed: string[] = [];
const helperFiles = new Set(Object.values(HELPERS));
for (const relative of new Bun.Glob("packages/*/src/**/*.{ts,tsx}").scanSync({
  cwd: REPO_ROOT,
})) {
  if (!isSourceFile(relative) || helperFiles.has(relative)) continue;
  const original = readFileSync(path.join(REPO_ROOT, relative), "utf8");
  if (!/Math\.(?:max|min)\(/u.test(original)) continue;
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
