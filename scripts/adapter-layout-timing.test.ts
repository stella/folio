import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import type { ObjectExpression } from "oxc-parser";

import { TRANSACTION_LAYOUT_TIMING } from "../packages/core/src/controller/layoutScheduler";
import { parseSourceModule, walkAst } from "./lib/oxc-ast";

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
      const { program } = parseSourceModule(
        file,
        readFileSync(join(import.meta.dir, "..", file), "utf8"),
      );
      const importsTiming = program.body.some(
        (statement) =>
          statement.type === "ImportDeclaration" &&
          statement.source.value === SCHEDULER_MODULE &&
          statement.specifiers.some(
            (element) =>
              element.type === "ImportSpecifier" &&
              element.local.name === TIMING_NAME,
          ),
      );
      expect(importsTiming).toBe(true);

      const schedulerConfigs: ObjectExpression[] = [];
      walkAst(program, (node) => {
        const config =
          node.type === "CallExpression" ? node.arguments.at(0) : undefined;
        if (
          node.type === "CallExpression" &&
          node.callee.type === "Identifier" &&
          node.callee.name === "createLayoutScheduler" &&
          node.arguments.length === 1 &&
          config?.type === "ObjectExpression"
        ) {
          schedulerConfigs.push(config);
        }
      });
      expect(schedulerConfigs).toHaveLength(1);

      const properties = schedulerConfigs[0]?.properties ?? [];
      expect(
        properties.filter(
          (property) =>
            property.type === "SpreadElement" &&
            property.argument.type === "Identifier" &&
            property.argument.name === TIMING_NAME,
        ),
      ).toHaveLength(1);
      expect(
        properties.filter(
          (property) =>
            property.type === "Property" &&
            property.kind === "init" &&
            !property.method &&
            !property.computed &&
            property.key.type === "Identifier" &&
            TIMING_PROPERTIES.has(property.key.name),
        ),
      ).toHaveLength(0);
    });
  }
});
