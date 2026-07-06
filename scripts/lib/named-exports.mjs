/**
 * Extract top-level named exports from a TypeScript source file.
 *
 * Walks the AST via TypeScript's compiler API (already a devDep). Captures:
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

import { readFileSync, existsSync } from "node:fs";
import { dirname, resolve } from "node:path";
import ts from "typescript";

export function collectNamedExports(entryPath, visited = new Set()) {
  if (visited.has(entryPath)) return new Set();
  visited.add(entryPath);

  if (!existsSync(entryPath)) return new Set();

  const source = readFileSync(entryPath, "utf8");
  const sf = ts.createSourceFile(entryPath, source, ts.ScriptTarget.Latest, true);
  const exports = new Set();
  const dir = dirname(entryPath);

  for (const node of sf.statements) {
    if (ts.isVariableStatement(node) && hasExport(node)) {
      for (const decl of node.declarationList.declarations) {
        if (ts.isIdentifier(decl.name)) exports.add(decl.name.text);
      }
    } else if (
      (ts.isFunctionDeclaration(node) ||
        ts.isClassDeclaration(node) ||
        ts.isInterfaceDeclaration(node) ||
        ts.isTypeAliasDeclaration(node) ||
        ts.isEnumDeclaration(node)) &&
      hasExport(node) &&
      node.name
    ) {
      exports.add(node.name.text);
    } else if (ts.isExportAssignment(node)) {
      exports.add("default");
    } else if (ts.isExportDeclaration(node)) {
      if (node.exportClause && ts.isNamedExports(node.exportClause)) {
        // export { X, Y as Z }
        for (const el of node.exportClause.elements) {
          exports.add(el.name.text);
        }
      } else if (!node.exportClause && node.moduleSpecifier) {
        // export * from '...'
        const spec = node.moduleSpecifier.text;
        if (spec.startsWith(".")) {
          const target = resolveRelative(dir, spec);
          if (target) {
            for (const name of collectNamedExports(target, visited)) exports.add(name);
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

function hasExport(node) {
  return node.modifiers?.some((m) => m.kind === ts.SyntaxKind.ExportKeyword) ?? false;
}

function resolveRelative(fromDir, spec) {
  const base = resolve(fromDir, spec);
  for (const cand of [`${base}.ts`, `${base}.tsx`, `${base}/index.ts`, `${base}/index.tsx`]) {
    if (existsSync(cand)) return cand;
  }
  return null;
}
