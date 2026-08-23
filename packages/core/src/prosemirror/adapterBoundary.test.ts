import { describe, expect, test } from "bun:test";

const ADAPTER_BOUNDARY_FILES = [
  "src/prosemirror/conversion/fromProseDoc.ts",
  "src/layout-bridge/convert/toFlowBlocks.ts",
  "src/prosemirror/validation.ts",
] as const;

const PROPERTY_CHANGE_CONSUMER_FILES = [
  "src/prosemirror/commands/comments.ts",
  "src/prosemirror/extensions/features/ListExtension.ts",
  "src/prosemirror/revisionCarriers.ts",
] as const;

const FORBIDDEN_RAW_ATTR_PATTERNS = [
  {
    pattern: /\b(?:node|child|rowNode|cellNode|contentNode|sdtChild)\.attrs\s+as\b/u,
    reason: "Raw node attrs must be narrowed through the typed attr readers.",
  },
  {
    pattern: /\b(?:node|child|rowNode|cellNode|contentNode|sdtChild)\.attrs\[/u,
    reason: "Bracket attr reads bypass path-aware validation diagnostics.",
  },
  {
    pattern: /\bconst\s+attrs\s*=\s*(?:node|child|rowNode|cellNode|contentNode|sdtChild)\.attrs\b/u,
    reason: "Do not alias raw attrs at conversion/layout boundaries.",
  },
  {
    pattern: /\b(?:pmAttrs|attrs)\._propertyChanges\s+as\b/u,
    reason:
      "Read paragraph property changes through expectParagraphAttrs; do not revalidate raw values downstream.",
  },
] as const;

describe("ProseMirror adapter attr boundaries", () => {
  test("keep conversion and layout attr reads behind typed readers", async () => {
    const violations: string[] = [];

    for (const file of ADAPTER_BOUNDARY_FILES) {
      // oxlint-disable-next-line no-await-in-loop -- sequential test scan that accumulates violations in source order for a stable assertion
      const source = await Bun.file(file).text();
      const lines = source.split("\n");

      for (const [index, line] of lines.entries()) {
        for (const { pattern, reason } of FORBIDDEN_RAW_ATTR_PATTERNS) {
          if (!pattern.test(line)) {
            continue;
          }
          violations.push(`${file}:${index + 1}: ${reason}`);
        }
      }
    }

    expect(violations).toEqual([]);
  });

  test("keep property-change consumers behind canonical paragraph attrs", async () => {
    const violations: string[] = [];

    for (const file of PROPERTY_CHANGE_CONSUMER_FILES) {
      // oxlint-disable-next-line no-await-in-loop -- sequential source scan keeps diagnostics stable
      const source = await Bun.file(file).text();
      for (const [index, line] of source.split("\n").entries()) {
        if (
          /\b(?:node\.attrs(?:\["_propertyChanges"\]|\._propertyChanges)|attrs\._propertyChanges)\b/u.test(
            line,
          )
        ) {
          violations.push(`${file}:${index + 1}`);
        }
      }
    }

    expect(violations).toEqual([]);
  });
});
