import { describe, expect, test } from "bun:test";

import { findNormalizationWarningViolations } from "./lib/normalization-warnings";

const module_ = (path: string, source: string) => ({ path, source });

describe("normalisation warning coverage", () => {
  test("a normaliser that reaches for a warning code passes", () => {
    expect(
      findNormalizationWarningViolations([
        module_("a/fooNormalization.ts", "warn({ code: PARSE_WARNING_CODES.duplicateNoteId });"),
      ]),
    ).toEqual([]);
  });

  test("a normaliser that neither warns nor exempts itself fails", () => {
    const violations = findNormalizationWarningViolations([
      module_("a/fooNormalization.ts", "export const normalizeFoo = () => {};"),
    ]);
    expect(violations).toHaveLength(1);
    expect(violations[0]?.path).toBe("a/fooNormalization.ts");
  });

  test("a lossless normaliser passes on a stated reason", () => {
    expect(
      findNormalizationWarningViolations([
        module_(
          "a/fooNormalization.ts",
          "// PARSE-WARNING-EXEMPT: renumbers ids without dropping anything.\nexport const f = () => {};",
        ),
      ]),
    ).toEqual([]);
  });

  test("an exemption with no reason is not a decision", () => {
    const violations = findNormalizationWarningViolations([
      module_("a/fooNormalization.ts", "// PARSE-WARNING-EXEMPT:\nexport const f = () => {};"),
    ]);
    expect(violations).toHaveLength(1);
    expect(violations[0]?.detail).toContain("no reason");
  });
});
