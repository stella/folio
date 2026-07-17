import { describe, expect, test } from "bun:test";

import { DEFAULT_TAB_STOP_TWIPS, parseSettings } from "./settingsParser";

const SETTINGS_HEAD = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">`;
const SETTINGS_TAIL = `</w:settings>`;

function wrap(inner: string): string {
  return `${SETTINGS_HEAD}${inner}${SETTINGS_TAIL}`;
}

describe("parseSettings — w:defaultTabStop (§17.6.13)", () => {
  test("returns OOXML default when settings.xml is missing", () => {
    expect(parseSettings(null).defaultTabStop).toBe(DEFAULT_TAB_STOP_TWIPS);
  });

  test("returns OOXML default when w:defaultTabStop element is absent", () => {
    expect(parseSettings(wrap("")).defaultTabStop).toBe(DEFAULT_TAB_STOP_TWIPS);
  });

  test("parses w:defaultTabStop val attribute", () => {
    expect(parseSettings(wrap(`<w:defaultTabStop w:val="1440"/>`)).defaultTabStop).toBe(1440);
  });

  test("ignores non-positive values (Word never emits them)", () => {
    expect(parseSettings(wrap(`<w:defaultTabStop w:val="0"/>`)).defaultTabStop).toBe(
      DEFAULT_TAB_STOP_TWIPS,
    );
    expect(parseSettings(wrap(`<w:defaultTabStop w:val="-100"/>`)).defaultTabStop).toBe(
      DEFAULT_TAB_STOP_TWIPS,
    );
  });

  test("ignores non-numeric values", () => {
    expect(parseSettings(wrap(`<w:defaultTabStop w:val="banana"/>`)).defaultTabStop).toBe(
      DEFAULT_TAB_STOP_TWIPS,
    );
  });

  test("ignores values beyond Word's maximum margin (~22 inches)", () => {
    // 50000 twips ≈ 34.7 inches — past any plausible page width; reject.
    expect(parseSettings(wrap(`<w:defaultTabStop w:val="50000"/>`)).defaultTabStop).toBe(
      DEFAULT_TAB_STOP_TWIPS,
    );
  });
});

describe("parseSettings — w:evenAndOddHeaders (§17.10.1)", () => {
  test("absent flag leaves evenAndOddHeaders undefined", () => {
    expect(parseSettings(wrap("")).evenAndOddHeaders).toBeUndefined();
    expect(parseSettings(null).evenAndOddHeaders).toBeUndefined();
  });

  test("a bare element records the on state", () => {
    expect(parseSettings(wrap(`<w:evenAndOddHeaders/>`)).evenAndOddHeaders).toBe(true);
  });

  test('an explicit w:val="0" is treated as off', () => {
    expect(
      parseSettings(wrap(`<w:evenAndOddHeaders w:val="0"/>`)).evenAndOddHeaders,
    ).toBeUndefined();
  });
});

describe("parseSettings — w:mirrorMargins (§17.15.1.57)", () => {
  test("records only the enabled state", () => {
    expect(parseSettings(wrap(`<w:mirrorMargins/>`)).mirrorMargins).toBe(true);
    expect(parseSettings(wrap(`<w:mirrorMargins w:val="0"/>`)).mirrorMargins).toBeUndefined();
    expect(parseSettings(wrap("")).mirrorMargins).toBeUndefined();
  });
});

describe("parseSettings — w:themeFontLang (§17.15.1.88)", () => {
  test("reads eastAsia and bidi tags", () => {
    expect(
      parseSettings(wrap(`<w:themeFontLang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>`))
        .themeFontLang,
    ).toEqual({ eastAsia: "ja-JP", bidi: "ar-SA" });
  });

  test("keeps only the EastAsian tag when bidi is absent", () => {
    expect(
      parseSettings(wrap(`<w:themeFontLang w:val="en-US" w:eastAsia="ja-JP"/>`)).themeFontLang,
    ).toEqual({ eastAsia: "ja-JP" });
  });

  test("is undefined when only the primary (w:val) lang is present", () => {
    expect(parseSettings(wrap(`<w:themeFontLang w:val="en-US"/>`)).themeFontLang).toBeUndefined();
  });

  test("is undefined when the element is absent", () => {
    expect(parseSettings(wrap("")).themeFontLang).toBeUndefined();
    expect(parseSettings(null).themeFontLang).toBeUndefined();
  });
});

describe("parseSettings — document line-breaking rules", () => {
  test("reads language-scoped prohibited line-start and line-end characters", () => {
    const settings = parseSettings(
      wrap(`
        <w:noLineBreaksBefore w:lang="ja-JP" w:val="、。"/>
        <w:noLineBreaksAfter w:lang="ja-JP" w:val="（［"/>
      `),
    );

    expect(settings.lineBreakRules).toEqual({
      noLineBreaksBefore: { language: "ja-JP", characters: "、。" },
      noLineBreaksAfter: { language: "ja-JP", characters: "（［" },
    });
  });

  test("reads the legacy Ethiopic and Amharic breaking compatibility flag", () => {
    expect(
      parseSettings(wrap(`<w:compat><w:applyBreakingRules/></w:compat>`)).lineBreakRules,
    ).toEqual({ useLegacyEthiopicAmharicRules: true });
  });

  test("ignores disabled and incomplete line-breaking controls", () => {
    expect(
      parseSettings(
        wrap(`
          <w:noLineBreaksBefore w:lang="ja-JP"/>
          <w:compat><w:applyBreakingRules w:val="0"/></w:compat>
        `),
      ).lineBreakRules,
    ).toBeUndefined();
  });
});

describe("parseSettings — application compatibility generation", () => {
  test("parses the named compatibilityMode setting", () => {
    expect(
      parseSettings(
        wrap(
          `<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/></w:compat>`,
        ),
      ).compatibilityMode,
    ).toBe(14);
  });

  test("ignores unrelated and malformed compatibility settings", () => {
    expect(
      parseSettings(
        wrap(
          `<w:compat><w:compatSetting w:name="other" w:val="14"/><w:compatSetting w:name="compatibilityMode" w:val="invalid"/></w:compat>`,
        ),
      ).compatibilityMode,
    ).toBeUndefined();
  });
});

describe("parseSettings — document automatic hyphenation", () => {
  test("reads the Word hyphenation controls", () => {
    expect(
      parseSettings(
        wrap(`
          <w:autoHyphenation/>
          <w:doNotHyphenateCaps/>
          <w:consecutiveHyphenLimit w:val="2"/>
          <w:hyphenationZone w:val="360"/>
        `),
      ),
    ).toMatchObject({
      autoHyphenation: true,
      doNotHyphenateCaps: true,
      consecutiveHyphenLimit: 2,
      hyphenationZoneTwips: 360,
    });
  });

  test("preserves explicit disabled on/off values", () => {
    expect(
      parseSettings(
        wrap(`
          <w:autoHyphenation w:val="0"/>
          <w:doNotHyphenateCaps w:val="false"/>
        `),
      ),
    ).toMatchObject({
      autoHyphenation: false,
      doNotHyphenateCaps: false,
    });
  });

  test("ignores absent, malformed, negative, and out-of-range integer controls", () => {
    expect(parseSettings(wrap("")).autoHyphenation).toBeUndefined();
    const settings = parseSettings(
      wrap(`
        <w:consecutiveHyphenLimit w:val="-1"/>
        <w:hyphenationZone w:val="31681"/>
      `),
    );

    expect(settings.consecutiveHyphenLimit).toBeUndefined();
    expect(settings.hyphenationZoneTwips).toBeUndefined();
  });
});

describe("parseSettings — break-only paragraph placement", () => {
  test("records only an enabled splitPgBreakAndParaMark compatibility flag", () => {
    expect(
      parseSettings(wrap(`<w:compat><w:splitPgBreakAndParaMark/></w:compat>`))
        .splitPageBreakAndParagraphMark,
    ).toBe(true);
    expect(
      parseSettings(wrap(`<w:compat><w:splitPgBreakAndParaMark w:val="0"/></w:compat>`))
        .splitPageBreakAndParagraphMark,
    ).toBeUndefined();
    expect(parseSettings(wrap("")).splitPageBreakAndParagraphMark).toBeUndefined();
  });
});
