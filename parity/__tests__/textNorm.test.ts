import { describe, expect, test } from "bun:test";

import { normalizeLineText } from "../textNorm";

describe("normalizeLineText", () => {
  test("collapses long dot leaders for stable TOC matching", () => {
    expect(normalizeLineText("Definitions ................ 2")).toBe("Definitions … 2");
    expect(
      normalizeLineText("Definitions ................................................ 2"),
    ).toBe("Definitions … 2");
  });

  test("canonicalizes PDF visual order for RTL dot-leader entries", () => {
    expect(normalizeLineText("1 ........ ........ أ- عنوان عربي")).toBe("أ- عنوان عربي … 1");
    expect(normalizeLineText("أ- عنوان عربي ........ 1")).toBe("أ- عنوان عربي … 1");
    expect(normalizeLineText("1 ........ 𞤀𞤣𞤢𞤤")).toBe("𞤀𞤣𞤢𞤤 … 1");
  });

  test("does not reorder LTR or non-leader numeric text", () => {
    expect(normalizeLineText("1 ........ Definitions")).toBe("1 … Definitions");
    expect(normalizeLineText("1 عنوان عربي")).toBe("1 عنوان عربي");
    expect(normalizeLineText("1 ........ Terms وشروط")).toBe("1 … Terms وشروط");
    expect(normalizeLineText("- عنوان عربي 1")).toBe("- عنوان عربي 1");
    expect(normalizeLineText("- Definitions 1 ........ 2")).toBe("- Definitions 1 … 2");
    expect(normalizeLineText("- المبلغ 100 ........ 5")).toBe("- المبلغ 100 … 5");
    expect(normalizeLineText("- رابط https://example.com 2 ........ 7")).toBe(
      "- رابط https://example.com 2 … 7",
    );
  });

  test("normalizes Symbol-font copyright extraction noise", () => {
    expect(normalizeLineText("\uf0e3 Loan Market Association")).toBe("ã Loan Market Association");
  });

  test("normalizes legacy Symbol-font bullets", () => {
    expect(normalizeLineText("\uf0b7 First item")).toBe("• First item");
  });

  test("folds CJK radical aliases emitted by PDF font maps", () => {
    expect(normalizeLineText("⺟甲⼄丙丁")).toBe("母甲乙丙丁");
  });
});
