import { describe, expect, test } from "bun:test";

import { normalizeLineText, textSimilarity } from "../textNorm";

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

const referenceDistance = (a: readonly string[], b: readonly string[]): number => {
  let previous = Array.from({ length: b.length + 1 }, (_, index) => index);
  for (let i = 1; i <= a.length; i += 1) {
    const current = [i];
    for (let j = 1; j <= b.length; j += 1) {
      current[j] = Math.min(
        (previous[j] ?? 0) + 1,
        (current[j - 1] ?? 0) + 1,
        (previous[j - 1] ?? 0) + (a[i - 1] === b[j - 1] ? 0 : 1),
      );
    }
    previous = current;
  }
  return previous[b.length] ?? 0;
};

describe("textSimilarity", () => {
  test("preserves normalized Levenshtein scores across ASCII and BMP inputs", () => {
    const fragments = [
      "",
      "a",
      "b",
      "ab",
      "ba",
      "Novák",
      "Nováka",
      "Novak",
      "novák",
      ".+[]",
      "é",
      "e\u0301",
    ];
    for (const a of fragments) {
      for (const b of fragments) {
        for (const left of [a, a + a]) {
          const maxLength = Math.max(left.length, b.length);
          expect(textSimilarity(left, b)).toBe(
            maxLength === 0 ? 1 : 1 - referenceDistance(left.split(""), b.split("")) / maxLength,
          );
        }
      }
    }
  });

  test("uses a scalar distance and denominator across generated astral pairs", () => {
    const fragments = [
      "",
      "a",
      "😀",
      "😃",
      "a😀",
      "😀a",
      "𞤀",
      "e\u0301",
      "é",
      "\ud800",
      "\udc00",
      "\ufffd",
    ];
    for (const a of fragments) {
      for (const b of fragments) {
        const left = Array.from((a + b).toWellFormed());
        const right = Array.from((b + a).toWellFormed());
        const maxLength = Math.max(left.length, right.length);
        expect(textSimilarity(a + b, b + a)).toBe(
          maxLength === 0 ? 1 : 1 - referenceDistance(left, right) / maxLength,
        );
      }
    }
  });

  test("pins astral edit and denominator units", () => {
    expect(textSimilarity("😀", "😃")).toBe(0);
    expect(textSimilarity("a😀", "a😃")).toBe(0.5);
    expect(textSimilarity("😀", "")).toBe(0);
    expect(textSimilarity("😀", "ab")).toBe(0);
    expect(textSimilarity("😀a", "a")).toBe(0.5);
    expect(textSimilarity("\ud800", "\udc00")).toBe(1);
    expect(textSimilarity("\ud800a", "\ufffdb")).toBe(0.5);
  });
});
