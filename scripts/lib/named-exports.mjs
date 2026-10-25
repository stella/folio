/**
 * Extract top-level named exports from a TypeScript source file.
 *
 * Walks the AST via oxc-parser (already a devDep). Captures:
 *   - `export const|let|var X` / `export function X` / `export class X`
 *   - `export type X` / `export interface X` / `export enum X`
 *   - `export { X, Y as Z, type W }` (alias resolves to the exported name)
 *   - `export { default as X } from '...'`
 *   - `export default ...` -> "default"
 *   - `export * from '...'` -> recurse one level for relative specifiers,
 *     bounded by a visited set
 *
 * Limitations: does not follow re-exports across package boundaries; an
 * `export *` from a bare specifier returns a symbolic `<*from:spec>` marker so
 * the parity report flags asymmetry without false positives.
 *
 * Adapted from the upstream docx-editor parity infra for the folio fork.
 */

import { existsSync, readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { parseSync } from "oxc-parser";

// Declarations that bind their own name when exported.
const NAMED_DECLARATIONS = new Set([
  "ClassDeclaration",
  "FunctionDeclaration",
  "TSDeclareFunction",
  "TSEnumDeclaration",
  "TSInterfaceDeclaration",
  "TSTypeAliasDeclaration",
]);

const moduleExportName = (name) =>
  name.type === "Identifier" ? name.name : name.value;

export function collectNamedExports(entryPath, visited = new Set()) {
  if (visited.has(entryPath)) return new Set();
  visited.add(entryPath);

  if (!existsSync(entryPath)) return new Set();

  const source = readFileSync(entryPath, "utf8");
  const { program, errors } = parseSync(entryPath, source);
  // A half-read module would report exports it never saw as missing.
  if (errors.length > 0) {
    throw new Error(`${entryPath}: ${errors[0].message}`);
  }
  const exports = new Set();
  const dir = dirname(entryPath);

  for (const node of program.body) {
    if (node.type === "ExportNamedDeclaration" && node.declaration) {
      const { declaration } = node;
      if (declaration.type === "VariableDeclaration") {
        for (const decl of declaration.declarations) {
          if (decl.id.type === "Identifier") exports.add(decl.id.name);
        }
      } else if (NAMED_DECLARATIONS.has(declaration.type) && declaration.id) {
        exports.add(declaration.id.name);
      }
    } else if (node.type === "ExportNamedDeclaration") {
      // export { X, Y as Z }
      for (const el of node.specifiers) {
        exports.add(moduleExportName(el.exported));
      }
    } else if (node.type === "ExportDefaultDeclaration") {
      // `export default function Foo()` / `export default class Bar` binds as
      // `default`, not the declaration identifier. An anonymous default
      // function or class declaration binds no name.
      if (!NAMED_DECLARATIONS.has(node.declaration.type) || node.declaration.id)
        exports.add("default");
    } else if (node.type === "TSExportAssignment") {
      exports.add("default");
    } else if (node.type === "ExportAllDeclaration") {
      if (node.exported) {
        // export * as ns from '...'
        exports.add(moduleExportName(node.exported));
      } else {
        // export * from '...'
        const spec = node.source.value;
        if (spec.startsWith(".")) {
          const target = resolveRelative(dir, spec);
          if (target) {
            for (const name of collectNamedExports(target, visited))
              exports.add(name);
          } else {
            exports.add(`<*from:${spec}>`);
          }
        } else {
          exports.add(`<*from:${spec}>`);
        }
      }
    }
  }

  return exports;
}

function resolveRelative(fromDir, spec) {
  const base = resolve(fromDir, spec);
  for (const cand of [
    `${base}.ts`,
    `${base}.tsx`,
    `${base}/index.ts`,
    `${base}/index.tsx`,
  ]) {
    if (existsSync(cand)) return cand;
  }
  return null;
}
