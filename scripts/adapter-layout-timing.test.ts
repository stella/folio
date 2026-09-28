import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import ts from "typescript";
import { TRANSACTION_LAYOUT_TIMING } from "../packages/core/src/controller/layoutScheduler";

const ADAPTER_FILES = [
  "packages/react/src/paged-editor/PagedEditor.tsx",
  "packages/vue/src/composables/useDocxEditor.ts",
] as const;
const SCHEDULER_MODULE = "@stll/folio-core/controller/layoutScheduler";
const TIMING_NAME = "TRANSACTION_LAYOUT_TIMING";
const TIMING_PROPERTIES = new Set(["debounceMs", "maxDelayMs", "leadingFrame"]);

describe("adapter transaction layout timing", () => {
  test("uses the established interactive timing", () => {
    expect(TRANSACTION_LAYOUT_TIMING).toEqual({
      debounceMs: 32,
      maxDelayMs: 96,
      leadingFrame: true,
    });
  });

  for (const file of ADAPTER_FILES) {
    test(`${file} uses the core timing in its scheduler`, () => {
      const source = ts.createSourceFile(
        file,
        readFileSync(join(import.meta.dir, "..", file), "utf8"),
        ts.ScriptTarget.Latest,
        true,
      );
      const importsTiming = source.statements.some(
        (statement) =>
          ts.isImportDeclaration(statement) &&
          ts.isStringLiteral(statement.moduleSpecifier) &&
          statement.moduleSpecifier.text === SCHEDULER_MODULE &&
          statement.importClause?.namedBindings !== undefined &&
          ts.isNamedImports(statement.importClause.namedBindings) &&
          statement.importClause.namedBindings.elements.some(
            (element) => element.name.text === TIMING_NAME,
          ),
      );
      expect(importsTiming).toBe(true);

      const schedulerConfigs: ts.ObjectLiteralExpression[] = [];
      const visit = (node: ts.Node): void => {
        const config = ts.isCallExpression(node) ? node.arguments.at(0) : undefined;
        if (
          ts.isCallExpression(node) &&
          ts.isIdentifier(node.expression) &&
          node.expression.text === "createLayoutScheduler" &&
          node.arguments.length === 1 &&
          config !== undefined &&
          ts.isObjectLiteralExpression(config)
        ) {
          schedulerConfigs.push(config);
        }
        ts.forEachChild(node, visit);
      };
      visit(source);
      expect(schedulerConfigs).toHaveLength(1);

      const properties = schedulerConfigs[0]?.properties ?? [];
      expect(
        properties.filter(
          (property) =>
            ts.isSpreadAssignment(property) &&
            ts.isIdentifier(property.expression) &&
            property.expression.text === TIMING_NAME,
        ),
      ).toHaveLength(1);
      expect(
        properties.filter(
          (property) =>
            (ts.isPropertyAssignment(property) || ts.isShorthandPropertyAssignment(property)) &&
            ts.isIdentifier(property.name) &&
            TIMING_PROPERTIES.has(property.name.text),
        ),
      ).toHaveLength(0);
    });
  }
});
