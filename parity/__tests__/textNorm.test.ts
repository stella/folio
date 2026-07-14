import { describe, expect, test } from "bun:test";

import { normalizeLineText } from "../textNorm";

describe("normalizeLineText", () => {
  test("collapses long dot leaders for stable TOC matching", () => {
    expect(normalizeLineText("Definitions ................ 2")).toBe("Definitions … 2");
    expect(
      normalizeLineText("Definitions ................................................ 2"),
    ).toBe("Definitions … 2");
  });

  test("normalizes Symbol-font copyright extraction noise", () => {
    expect(normalizeLineText("\uf0e3 Loan Market Association")).toBe("ã Loan Market Association");
  });

  test("folds Kangxi radical aliases emitted by PDF font maps", () => {
    expect(normalizeLineText("甲⼄丙丁")).toBe("甲乙丙丁");
  });
});
