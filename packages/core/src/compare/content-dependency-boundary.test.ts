import { describe, expect, test } from "bun:test";

const NEUTRAL_COMPARISON_MODULES = [
  "column-alignment.ts",
  "content-alignment.ts",
  "content-properties.ts",
  "content-types.ts",
  "content.ts",
  "formatting.ts",
  "text-diff.ts",
] as const;

const FORBIDDEN_IMPORT_PREFIXES = [
  "../ai-edits",
  "../docx",
  "../prosemirror",
  "@stll/docx",
  "prosemirror-",
] as const;

describe("neutral comparison dependency boundary", () => {
  test("the semantic core imports only representation-neutral modules", async () => {
    for (const moduleName of NEUTRAL_COMPARISON_MODULES) {
      const source = await Bun.file(`${import.meta.dir}/${moduleName}`).text();
      const imports = [...source.matchAll(/(?:\bfrom\s*|\bimport\s*)["']([^"']+)["']/gu)].map(
        (match) => match[1] ?? "",
      );
      expect(
        imports.filter((specifier) =>
          FORBIDDEN_IMPORT_PREFIXES.some((prefix) => specifier.startsWith(prefix)),
        ),
        moduleName,
      ).toEqual([]);
    }
  });
});
