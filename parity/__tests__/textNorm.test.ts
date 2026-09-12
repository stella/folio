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

  test("normalizes legacy Wingdings checkboxes", () => {
    expect(normalizeLineText("\uf0a8 Apply")).toBe("☐ Apply");
  });

  test("folds Persian code points a PDF font map returns for Arabic letters", () => {
    expect(
      normalizeLineText(
        "\u0648\u0632\u0627\u0631\u0629 \u0627\u0644\u062a\u0639\u0644\u06cc\u0645",
      ),
    ).toBe(
      normalizeLineText(
        "\u0648\u0632\u0627\u0631\u0629 \u0627\u0644\u062a\u0639\u0644\u064a\u0645",
      ),
    );
    expect(normalizeLineText("\u0627\u0644\u0645\u0642\u0628\u0648\u0644\u0629 \u06be\u064a")).toBe(
      normalizeLineText("\u0627\u0644\u0645\u0642\u0628\u0648\u0644\u0629 \u0647\u064a"),
    );
  });

  test("folds mirrored bracket pairs on an RTL line", () => {
    expect(normalizeLineText(")\u0633\u0628\u0628 \u0627\u0644\u063a\u064a\u0627\u0628(")).toBe(
      normalizeLineText("(\u0633\u0628\u0628 \u0627\u0644\u063a\u064a\u0627\u0628)"),
    );
  });

  test("keeps bracket direction on an LTR line", () => {
    expect(normalizeLineText(")Reason for absence(")).not.toBe(
      normalizeLineText("(Reason for absence)"),
    );
  });

  test("folds CJK radical aliases emitted by PDF font maps", () => {
    expect(normalizeLineText("⺟甲⼄丙丁")).toBe("母甲乙丙丁");
  });
});
