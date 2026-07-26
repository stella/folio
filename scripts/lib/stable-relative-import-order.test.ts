import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import coreConfigs from "../../packages/core/tsdown.config";
import agentsConfigs from "../../packages/agents/tsdown.config";
import {
  canonicalizeRelativeImportOrder,
  STABLE_RELATIVE_IMPORT_ORDER_PLUGIN,
} from "./stable-relative-import-order";

const RELATIVE_IMPORTS = [
  'import { alpha } from "./alpha.js";',
  'import { beta } from "../beta.js";',
  'import "./side-effect-free.js";',
  'import { zeta } from "./zeta.js";',
] as const;

const firstGroupPermutation = fc.uniqueArray(fc.integer({ min: 0, max: 1 }), {
  minLength: 2,
  maxLength: 2,
});
const secondGroupPermutation = fc.uniqueArray(fc.integer({ min: 2, max: 3 }), {
  minLength: 2,
  maxLength: 2,
});

const importSource = (line: string) => line.match(/"([^"]+)"/)?.at(1);

const sortImports = (imports: readonly string[]) =>
  [...imports].sort((left, right) => {
    const leftSource = importSource(left) ?? "";
    const rightSource = importSource(right) ?? "";
    if (leftSource !== rightSource) {
      return leftSource < rightSource ? -1 : 1;
    }
    if (left === right) {
      return 0;
    }
    return left < right ? -1 : 1;
  });

const relativeImportSpans = (lines: readonly string[]) => {
  let offset = 0;
  let group = 0;
  let previousWasRelative = false;
  const spans: { start: number; end: number; source: string; group: number }[] = [];
  for (const line of lines) {
    const source = importSource(line);
    if (source?.startsWith(".") && (line.startsWith('import "') || line.includes(' from ".'))) {
      if (!previousWasRelative) {
        group += 1;
      }
      spans.push({ start: offset, end: offset + line.length, source, group });
      previousWasRelative = true;
    } else {
      previousWasRelative = false;
    }
    offset += line.length + 1;
  }
  return spans;
};

describe("stableRelativeImportOrder", () => {
  test("canonicalizes every relative-import permutation without moving external imports", () => {
    fc.assert(
      fc.property(firstGroupPermutation, secondGroupPermutation, (first, second) => {
        const firstGroup = first.flatMap((index) => {
          const line = RELATIVE_IMPORTS.at(index);
          return line ? [line] : [];
        });
        const secondGroup = second.flatMap((index) => {
          const line = RELATIVE_IMPORTS.at(index);
          return line ? [line] : [];
        });
        const lines = [
          ...firstGroup,
          'import { externalA } from "external-a";',
          ...secondGroup,
          'import "external-b";',
          "export {};",
        ];
        const expected = [
          ...sortImports(firstGroup),
          'import { externalA } from "external-a";',
          ...sortImports(secondGroup),
          'import "external-b";',
          "export {};",
        ].join("\n");
        const code = lines.join("\n");
        const canonical = canonicalizeRelativeImportOrder(code, relativeImportSpans(lines));

        expect(canonical).toBe(expected);
        expect(
          canonicalizeRelativeImportOrder(canonical, relativeImportSpans(expected.split("\n"))),
        ).toBe(canonical);
      }),
    );
  });

  test("source-mirrored JavaScript builds install the canonicalizer", () => {
    for (const config of [coreConfigs.at(0), agentsConfigs.at(0)]) {
      expect(config?.plugins).toEqual(
        expect.arrayContaining([
          expect.objectContaining({ name: STABLE_RELATIVE_IMPORT_ORDER_PLUGIN }),
        ]),
      );
    }
  });
});
