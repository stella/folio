import { describe, expect, test } from "bun:test";

import type { ParagraphAttrs, TextRun } from "../types";
import {
  lineBreakPolicyCacheParts,
  resolveEffectiveLineBreakPolicy,
} from "./effectiveLineBreakPolicy";

const textRun = (text: string, formatting: Partial<TextRun> = {}): TextRun => ({
  kind: "text",
  text,
  ...formatting,
});

const localeCases: [label: string, run: TextRun, locale: string][] = [
  [
    "Latin",
    textRun("text", { language: { val: "en-US", eastAsia: "ja-JP", bidi: "ar-SA" } }),
    "en-US",
  ],
  [
    "East Asian",
    textRun("日本", { language: { val: "en-US", eastAsia: "ja-JP", bidi: "ar-SA" } }),
    "ja-JP",
  ],
  [
    "bidirectional",
    textRun("نص", { rtl: true, language: { val: "en-US", eastAsia: "ja-JP", bidi: "ar-SA" } }),
    "ar-SA",
  ],
];

describe("resolveEffectiveLineBreakPolicy", () => {
  test.each(localeCases)("selects the %s run language", (_label, run, locale) => {
    expect(resolveEffectiveLineBreakPolicy({ attrs: undefined, run }).provider.locale).toBe(locale);
  });

  test("applies only custom line-edge lists matching the active language", () => {
    const attrs: ParagraphAttrs = {
      kinsoku: true,
      lineBreakRules: {
        noLineBreaksBefore: { language: "ja", characters: "。" },
        noLineBreaksAfter: { language: "zh-CN", characters: "（" },
        useLegacyEthiopicAmharicRules: true,
      },
    };

    expect(
      resolveEffectiveLineBreakPolicy({
        attrs,
        run: textRun("日本", { language: { val: "en-US", eastAsia: "ja-JP" } }),
      }).provider,
    ).toEqual({
      locale: "ja-JP",
      kinsoku: true,
      noLineBreaksBefore: "。",
      useLegacyEthiopicAmharicRules: true,
    });
  });

  test("keeps omitted provider values omitted while resolving layout defaults", () => {
    expect(resolveEffectiveLineBreakPolicy({ attrs: undefined, run: textRun("plain") })).toEqual({
      provider: {},
      hangingPunctuation: true,
      automaticHyphenation: { type: "disabled" },
    });
  });

  test("resolves explicit punctuation and automatic-hyphenation state", () => {
    const attrs: ParagraphAttrs = {
      overflowPunctuation: false,
      automaticHyphenation: {
        enabled: true,
        doNotHyphenateCaps: true,
      },
    };

    expect(
      resolveEffectiveLineBreakPolicy({
        attrs,
        run: textRun("legal", { allCaps: true, language: { val: "en-US" } }),
      }),
    ).toEqual({
      provider: {
        locale: "en-US",
        doNotHyphenateCaps: true,
        renderedAllCaps: true,
      },
      hangingPunctuation: false,
      automaticHyphenation: {
        type: "enabled",
        consecutiveLineLimit: 0,
        hyphenationZoneTwips: 360,
      },
    });
  });

  test("paragraph suppression disables hyphenation without changing provider inputs", () => {
    const attrs: ParagraphAttrs = {
      suppressAutoHyphens: true,
      automaticHyphenation: {
        enabled: true,
        doNotHyphenateCaps: true,
        consecutiveLineLimit: 2,
        hyphenationZoneTwips: 720,
      },
    };
    const original = structuredClone(attrs);

    const policy = resolveEffectiveLineBreakPolicy({ attrs, run: textRun("text") });

    expect(policy.automaticHyphenation).toEqual({ type: "disabled" });
    expect(policy.provider.doNotHyphenateCaps).toBe(true);
    expect(attrs).toEqual(original);
  });
});

describe("lineBreakPolicyCacheParts", () => {
  test("covers every authored paragraph-level policy input", () => {
    expect(
      lineBreakPolicyCacheParts({
        kinsoku: false,
        overflowPunctuation: true,
        suppressAutoHyphens: false,
        automaticHyphenation: {
          enabled: true,
          doNotHyphenateCaps: true,
          consecutiveLineLimit: 2,
          hyphenationZoneTwips: 720,
        },
        lineBreakRules: {
          noLineBreaksBefore: { language: "ja", characters: "。" },
          noLineBreaksAfter: { language: "ja", characters: "（" },
          useLegacyEthiopicAmharicRules: true,
        },
      }),
    ).toEqual([
      "kinsoku:false",
      "overflow-punct:true",
      "suppress-auto-hyphens:false",
      "auto-hyphens:true|2|720",
      "line-break-rules:ja|。|ja|（|true",
    ]);
  });
});
