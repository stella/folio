import { describe, expect, test } from "bun:test";

import { FOLIO_LOCALES, getFolioMessages, isFolioLocale } from "./messages";

// Toolbar labels folio added after the initial translated catalogs were
// captured. They are now translated in every non-English locale; the
// i18n-check gate (scripts/i18n-check.ts) stops them regressing to the English
// source.
const POST_EXTRACTION_KEYS = [
  "insertGroup",
  "insertImage",
  "insertPageBreak",
  "insertTable",
  "insertTableOfContents",
  "zoomGroup",
] as const;

const flatten = (obj: Record<string, unknown>, prefix = ""): Map<string, string> => {
  const out = new Map<string, string>();
  for (const [key, value] of Object.entries(obj)) {
    const path = prefix ? `${prefix}.${key}` : key;
    if (typeof value === "string") {
      out.set(path, value);
      continue;
    }
    if (value && typeof value === "object") {
      for (const [k, v] of flatten(value as Record<string, unknown>, path)) {
        out.set(k, v);
      }
    }
  }
  return out;
};

const folioOf = (locale: string): Map<string, string> =>
  flatten(getFolioMessages(locale).folio as Record<string, unknown>);

const englishKeys = folioOf("en");

describe("getFolioMessages", () => {
  test("ships exactly the 17 documented locales", () => {
    expect([...FOLIO_LOCALES]).toEqual([
      "en",
      "de",
      "fr",
      "es",
      "cs",
      "ar",
      "et",
      "he",
      "hi",
      "hu",
      "lt",
      "lv",
      "pl",
      "pt-BR",
      "sk",
      "tr",
      "zh-CN",
    ]);
    expect(new Set(FOLIO_LOCALES).size).toBe(FOLIO_LOCALES.length);
  });

  test("isFolioLocale narrows to bundled locales only", () => {
    expect(isFolioLocale("de")).toBe(true);
    expect(isFolioLocale("pt-BR")).toBe(true);
    expect(isFolioLocale("xx")).toBe(false);
    expect(isFolioLocale("en-US")).toBe(false);
  });

  test("falls back to English for an unbundled locale", () => {
    expect(getFolioMessages("xx")).toBe(getFolioMessages("en"));
  });

  test("resolves a regional tag to its base-language catalog", () => {
    expect(getFolioMessages("de-DE")).toBe(getFolioMessages("de"));
    expect(getFolioMessages("fr-FR")).toBe(getFolioMessages("fr"));
    // pt-BR ships in its own right; pt-PT has no base match, so it falls back.
    expect(getFolioMessages("pt-PT")).toBe(getFolioMessages("en"));
  });

  test("falls back to English for a structurally invalid tag", () => {
    expect(getFolioMessages("!!invalid")).toBe(getFolioMessages("en"));
  });

  test("English is 100% complete (every key present and non-empty)", () => {
    expect(englishKeys.size).toBeGreaterThan(0);
    for (const [key, value] of englishKeys) {
      expect(value.length, `folio.${key} must be a non-empty English string`).toBeGreaterThan(0);
    }
  });

  test("every locale has the exact same key set as English (no missing/extra keys)", () => {
    const expected = [...englishKeys.keys()].sort();
    for (const locale of FOLIO_LOCALES) {
      const actual = [...folioOf(locale).keys()].sort();
      expect(actual, `locale ${locale} key set`).toEqual(expected);
    }
  });

  test("no translated value is empty in any locale", () => {
    for (const locale of FOLIO_LOCALES) {
      for (const [key, value] of folioOf(locale)) {
        expect(value.length, `${locale}: folio.${key} is empty`).toBeGreaterThan(0);
      }
    }
  });

  test("a toolbar label renders translated, not the English source, in non-English locales", () => {
    // `folio.bold` is a stable toolbar aria-label (FormattingBar).
    expect(folioOf("de").get("bold")).toBe("Fett");
    expect(folioOf("cs").get("bold")).toBe("Tučné");
    expect(folioOf("ar").get("bold")).toBe("عريض");
    expect(folioOf("ar").get("bold")).toMatch(/\p{Script=Arabic}/u);
    expect(folioOf("pt-BR").get("bold")).toBe("Negrito");
    expect(folioOf("zh-CN").get("bold")).toBe("加粗");

    const english = englishKeys.get("bold");
    for (const locale of FOLIO_LOCALES) {
      if (locale === "en") {
        continue;
      }
      expect(folioOf(locale).get("bold"), `${locale} should translate folio.bold`).not.toBe(
        english,
      );
    }
  });

  test("the common date-picker label is folded into folio.* and translated", () => {
    expect(folioOf("en").get("clearDate")).toBe("Clear date");
    expect(folioOf("de").get("clearDate")).toBe("Datum löschen");
    expect(folioOf("ar").get("clearDate")).toMatch(/\p{Script=Arabic}/u);
    expect(folioOf("pt-BR").get("clearDate")).toBe("Limpar data");
    expect(folioOf("zh-CN").get("clearDate")).toBe("清除日期");
  });

  test("the post-extraction keys are translated (not the English source) in every non-English locale", () => {
    for (const key of POST_EXTRACTION_KEYS) {
      const english = englishKeys.get(key);
      expect(english, `folio.${key} must exist in English`).toBeDefined();
      for (const locale of FOLIO_LOCALES) {
        if (locale === "en") {
          continue;
        }
        expect(folioOf(locale).get(key), `${locale}: folio.${key} should be translated`).not.toBe(
          english,
        );
      }
    }
  });
});
