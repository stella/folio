import { describe, expect, test } from "bun:test";

import {
  getLineBreakProvider,
  resetLineBreakProvider,
  setLineBreakProvider,
} from "./lineBreakProvider";
import {
  findGraphemeBreaks,
  findHyphenationBreaks,
  findWordBreaks,
  isHangingPunctuation,
  isBreakChar,
} from "./lineBreaks";

describe("findWordBreaks", () => {
  test("breaks after spaces and hyphens in Latin text", () => {
    expect(findWordBreaks("hello world")).toEqual([6]);
    expect(findWordBreaks("well-known")).toEqual([5]);
    expect(findWordBreaks("a  b")).toEqual([2, 3]);
    expect(findWordBreaks("hy\u00ADphen")).toEqual([3]);
  });

  test("treats an ASCII CRLF pair as one grapheme", () => {
    expect(findWordBreaks("first\r\nsecond")).toEqual([7]);
  });

  test("finds breaks in precomposed international Latin text", () => {
    expect(findWordBreaks("Příliš žluťoučký kůň úpěl ďábelské ódy.", { locale: "cs-CZ" })).toEqual([
      7, 17, 21, 26, 35,
    ]);
  });

  test("adds a break after every CJK code point", () => {
    expect(findWordBreaks("日本語")).toEqual([1, 2, 3]);
    expect(findWordBreaks("A世")).toEqual([2]);
  });

  test("handles astral CJK ideographs as single break units", () => {
    const text = "𠀀文";
    const breaks = findWordBreaks(text);
    expect(breaks).toEqual([2, 3]);
  });

  test("uses dictionary boundaries for scripts without spaces", () => {
    expect(findWordBreaks("ภาษาไทยภาษาไทย", { locale: "th" })).toEqual([4, 7, 11]);
  });

  test("falls back safely for malformed DOCX language tags", () => {
    expect(() => findWordBreaks("hello world", { locale: "not_a_locale" })).not.toThrow();
  });

  test("keeps ordinary spaces breakable regardless of locale", () => {
    const text = "text v  text";

    expect(findWordBreaks(text, { locale: "cs-CZ" })).toEqual([5, 7, 8]);
    expect(findWordBreaks(text, { locale: "en-US" })).toEqual([5, 7, 8]);
    expect(findWordBreaks("textov text", { locale: "cs-CZ" })).toEqual([7]);
    expect(findWordBreaks("text o text", { locale: "cs-CZ" })).toEqual([5, 7]);
    expect(findWordBreaks("text a text", { locale: "cs-CZ" })).toEqual([5, 7]);
  });

  test("does not break at glue-class spaces", () => {
    expect(findWordBreaks("one\u00A0two")).toEqual([]);
    expect(findWordBreaks("one\u2007two")).toEqual([]);
    expect(findWordBreaks("one\u202Ftwo")).toEqual([]);
  });

  test("keeps prohibited CJK punctuation off the next line", () => {
    expect(findWordBreaks("中文。测试", { locale: "zh-CN" })).toEqual([1, 3, 4, 5]);
  });

  test("applies document-specific kinsoku overrides", () => {
    expect(findWordBreaks("中文測", { noLineBreaksBefore: new Set(["測"]) })).toEqual([1, 3]);
  });

  test("ignores custom kinsoku lists when kinsoku is disabled", () => {
    expect(
      findWordBreaks("中文測", {
        kinsoku: false,
        noLineBreaksBefore: new Set(["測"]),
      }),
    ).toEqual([1, 2, 3]);
  });

  test("replaces the default kinsoku lists with language-specific settings", () => {
    expect(
      findWordBreaks("中文。」測", {
        locale: "ja-JP",
        kinsoku: true,
        noLineBreaksBefore: new Set(["。"]),
      }),
    ).toEqual([1, 3, 4, 5]);
    expect(
      findWordBreaks("中（文", {
        locale: "ja-JP",
        kinsoku: true,
        noLineBreaksAfter: new Set(["測"]),
      }),
    ).toEqual([1, 2, 3]);
  });

  test("applies legacy Ethiopic and Amharic break opportunities", () => {
    const text = "ሀ፡ለ";
    expect(findWordBreaks(text)).toEqual([]);
    expect(findWordBreaks(text, { useLegacyEthiopicAmharicRules: true })).toEqual([2]);
  });
});

describe("findGraphemeBreaks", () => {
  test("finds simple precomposed graphemes and keeps CRLF together", () => {
    expect(findGraphemeBreaks("café\r\nx")).toEqual([1, 2, 3, 4, 6, 7]);
  });

  test("never splits an emoji ZWJ sequence or combining character", () => {
    expect(findGraphemeBreaks("👨‍👩‍👧‍👦x")).toEqual([11, 12]);
    expect(findGraphemeBreaks("e\u0301x")).toEqual([2, 3]);
  });
});

describe("findHyphenationBreaks", () => {
  test("uses the dictionary selected by the exact Word language", () => {
    expect(findHyphenationBreaks("hyphenation", { locale: "en-US" })).toEqual([2, 6]);
    expect(findHyphenationBreaks("nejneobhospodářovávatelnější", { locale: "cs-CZ" })).toEqual([
      3, 5, 7, 10, 12, 14, 16, 18, 20, 23, 26,
    ]);
  });

  test("fails closed for unsupported or missing locales", () => {
    expect(findHyphenationBreaks("hyphenation")).toEqual([]);
    expect(findHyphenationBreaks("hyphenation", { locale: "en" })).toEqual([]);
    expect(findHyphenationBreaks("hyphenation", { locale: "fr-FR" })).toEqual([]);
    expect(() =>
      findHyphenationBreaks("HYPHENATION", {
        locale: "en-US-%%%",
        doNotHyphenateCaps: true,
      }),
    ).not.toThrow();
  });

  test("bounds dictionary work for hostile oversized words", () => {
    expect(findHyphenationBreaks("a".repeat(257), { locale: "en-US" })).toEqual([]);
  });

  test("honors the document all-caps setting", () => {
    expect(findHyphenationBreaks("HYPHENATION", { locale: "en-US" })).not.toEqual([]);
    expect(
      findHyphenationBreaks("HYPHENATION", {
        locale: "en-US",
        doNotHyphenateCaps: true,
      }),
    ).toEqual([]);
    expect(
      findHyphenationBreaks("hyphenation", {
        locale: "en-US",
        doNotHyphenateCaps: true,
        renderedAllCaps: true,
      }),
    ).toEqual([]);
  });
});

describe("isHangingPunctuation", () => {
  test("allows closing punctuation but not opening punctuation to overhang", () => {
    expect(isHangingPunctuation("。", { locale: "zh-CN" })).toBe(true);
    expect(isHangingPunctuation("（", { locale: "zh-CN" })).toBe(false);
  });

  test("does not apply East-Asian hanging punctuation to Latin runs", () => {
    expect(isHangingPunctuation(",", { locale: "en-US" })).toBe(false);
    expect(isHangingPunctuation(",", { locale: "ja-JP" })).toBe(true);
  });

  test("includes document-specific prohibited line-start characters", () => {
    expect(isHangingPunctuation("※", { locale: "ja-JP", noLineBreaksBefore: new Set(["※"]) })).toBe(
      true,
    );
    expect(
      isHangingPunctuation("※", {
        locale: "ja-JP",
        kinsoku: false,
        noLineBreaksBefore: new Set(["※"]),
      }),
    ).toBe(false);
  });

  test("does not hang default punctuation omitted from a custom list", () => {
    expect(
      isHangingPunctuation("」", {
        locale: "ja-JP",
        kinsoku: true,
        noLineBreaksBefore: new Set(["。"]),
      }),
    ).toBe(false);
  });
});

describe("line-break provider", () => {
  test("is replaceable and resettable", () => {
    const original = getLineBreakProvider();
    try {
      setLineBreakProvider({
        findBreaks: () => [1],
        findGraphemeBreaks: () => [1],
      });
      expect(findWordBreaks("abc")).toEqual([1]);
    } finally {
      resetLineBreakProvider();
    }
    expect(getLineBreakProvider()).toBe(original);
  });
});

describe("isBreakChar", () => {
  test("treats CJK code points and low surrogates as break chars", () => {
    expect(isBreakChar("世")).toBe(true);
    expect(isBreakChar("A")).toBe(false);
    expect(isBreakChar("\uDC00")).toBe(true);
  });
});
