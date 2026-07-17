import { describe, expect, test } from "bun:test";

import { getResolvedData } from "../layout-engine/measure/measureHelpers";
import {
  CJK_FALLBACK_FONT_FAMILY,
  getGoogleFontEquivalent,
  getGoogleFontsEnabled,
  getGoogleFontsToLoad,
  hasGoogleFontEquivalent,
  isCjkFont,
  resolveFontFamily,
  setGoogleFontsEnabled,
} from "./fontResolver";

describe("fontResolver — single-line ratios are derived from real hhea metrics", () => {
  // Expected ratios computed by hand from each font's real `hhea` table:
  // (hheaAscent + |hheaDescent| + hheaLineGap) / unitsPerEm. Measured against
  // real Word rendering (11pt, single spacing) for every font in this table.
  // This locks the facts driving FONT_MAPPINGS — a future edit that
  // hand-writes a different decimal into the table will fail here.
  const verifiedRatios: [font: string, expectedRatio: number][] = [
    ["aptos", 1.2207],
    ["aptos display", 1.2207],
    ["arial", 1.1499],
    ["times new roman", 1.1499],
    ["cg times", 1.1499],
    ["calibri", 1.2207],
    ["cambria", 1.1724],
    ["georgia", 1.1362],
    ["verdana", 1.2153],
    ["tahoma", 1.207],
    ["segoe ui", 1.3301],
    ["montserrat", 1.219],
    ["trebuchet ms", 1.1611],
    ["courier new", 1.1328],
    ["consolas", 1.1709],
    ["comic sans ms", 1.3936],
    ["impact", 1.2197],
    ["palatino linotype", 1.3491],
    ["book antiqua", 1.2056],
    ["garamond", 1.1273],
    ["century gothic", 1.2261],
    ["lucida console", 1.0],
  ];

  for (const [font, expectedRatio] of verifiedRatios) {
    test(`${font} → ${expectedRatio}`, () => {
      expect(resolveFontFamily(font).singleLineRatio).toBeCloseTo(expectedRatio, 4);
    });
  }
});

describe("fontResolver — Aptos falls back to bundled Lato", () => {
  test.each(["Aptos", "Aptos Display"])("%s keeps an Office-compatible fallback", (font) => {
    const resolved = resolveFontFamily(font);

    expect(resolved.googleFont).toBe("Lato");
    expect(resolved.cssFallback).toContain("Lato");
    expect(resolved.hasGoogleEquivalent).toBe(true);
  });
});

describe("fontResolver — legacy Times-compatible faces", () => {
  test("CG Times uses Times New Roman metrics and CSS fallback", () => {
    const resolved = resolveFontFamily("CG Times");

    expect(resolved.originalFont).toBe("CG Times");
    expect(resolved.googleFont).toBe("Tinos");
    expect(resolved.cssFallback.startsWith('"Times New Roman", Tinos')).toBe(true);
    expect(getResolvedData("CG Times").singleLineRatio).toBeCloseTo(1.1499, 4);
  });
});

describe("fontResolver — previously wrong ratios are corrected and reach consumers", () => {
  // These fonts hand-transcribed a ratio that dropped or mis-stated the line
  // gap; the corrected values come from the real hhea metrics above. Assert
  // through `getResolvedData` (the layout engine's public resolution path,
  // also used by `measureContainer.ts`) to prove the fix reaches consumers,
  // not just the raw `resolveFontFamily` call.
  const correctedCases: [font: string, correctedRatio: number][] = [
    ["cambria", 1.1724], // was hand-transcribed 1.2676 — wrong by 8%
    ["palatino linotype", 1.3491], // was hand-transcribed 1.0259 — wrong by 31%
    ["arial", 1.1499], // fixed in a prior revision; still exercised here
    ["book antiqua", 1.2056], // was hand-transcribed 1.0259 — wrong by 17%
    ["garamond", 1.1273], // replaces the unverified 1.068 legacy value
    ["century gothic", 1.2261], // was hand-transcribed 1.1611 — wrong by 6%
    ["trebuchet ms", 1.1611], // was hand-transcribed 1.1431 — wrong
    ["consolas", 1.1709], // was hand-transcribed 1.1626 — wrong
    ["lucida console", 1.0], // was hand-transcribed 1.1387 — wrong by 14%
  ];

  for (const [font, correctedRatio] of correctedCases) {
    test(`${font} resolves to the corrected ratio via getResolvedData`, () => {
      expect(getResolvedData(font).singleLineRatio).toBeCloseTo(correctedRatio, 4);
    });
  }
});

describe("fontResolver — unverified legacy fonts are left unchanged", () => {
  // The original face is not available here to read hhea metrics from, so its
  // hand-transcribed ratio stays unchanged until a verified source is available.
  const legacyCases: [font: string, unchangedRatio: number][] = [["lucida sans", 1.1655]];

  for (const [font, unchangedRatio] of legacyCases) {
    test(`${font} keeps its legacy ratio`, () => {
      expect(resolveFontFamily(font).singleLineRatio).toBeCloseTo(unchangedRatio, 4);
    });
  }
});

describe("fontResolver — Japanese Mincho/Gothic carry the measured East-Asian line height", () => {
  // Word font-links CJK glyphs to the default East-Asian face and uses its
  // taller single-line height; measured against real Word on a non-grid
  // Japanese document (10.5pt body → 13.68pt line pitch ≈ 1.303×). The value
  // is approximate for the CJK long tail (Yu Mincho renders ≈1.60 when
  // actually installed); these families all share the MS Mincho/Gothic default.
  const measuredJpFaces = [
    "MS Mincho",
    "MS Gothic",
    "ＭＳ 明朝",
    "ＭＳ Ｐ明朝",
    "ＭＳ ゴシック",
    "ＭＳ Ｐゴシック",
    // Romanized/native aliases sharing the MS Mincho/Gothic entries.
    "Yu Mincho",
    "Yu Gothic",
    "游明朝",
    "游ゴシック",
    "Meiryo",
  ];

  for (const face of measuredJpFaces) {
    test(`${face} → 1.303`, () => {
      expect(resolveFontFamily(face).singleLineRatio).toBeCloseTo(1.303, 4);
    });
  }

  test("the CJK line-height fallback family resolves to the measured ratio via getResolvedData", () => {
    // measureParagraph resolves the fallback through getResolvedData, so the
    // constant must land on the measured 1.303 entry through that path too.
    expect(getResolvedData(CJK_FALLBACK_FONT_FAMILY).singleLineRatio).toBeCloseTo(1.303, 4);
  });

  test("the CJK long tail keeps the default ratio until measured", () => {
    for (const face of ["SimSun", "SimHei", "宋体", "Malgun Gothic", "바탕"]) {
      expect(resolveFontFamily(face).singleLineRatio).toBeCloseTo(1.15, 4);
    }
  });
});

describe("fontResolver — isCjkFont classifies East-Asian faces", () => {
  // True only for mapped CJK entries (direct or via romanized alias); Latin
  // faces and unmapped families are false so the measurer falls back to
  // CJK_FALLBACK_FONT_FAMILY for them.
  const cjkFaces = [
    "MS Mincho",
    "ms gothic",
    "ＭＳ Ｐ明朝",
    "SimSun",
    "宋体",
    "微軟正黑體",
    "Malgun Gothic",
    "굴림",
    "Yu Mincho",
    "Meiryo",
  ];
  const nonCjkFaces = [
    "Arial",
    "Century",
    "Calibri",
    "Times New Roman",
    "Garamond",
    "Unknown Face",
  ];

  for (const face of cjkFaces) {
    test(`${face} is CJK`, () => {
      expect(isCjkFont(face)).toBe(true);
    });
  }

  for (const face of nonCjkFaces) {
    test(`${face} is not CJK`, () => {
      expect(isCjkFont(face)).toBe(false);
    });
  }
});

describe("fontResolver — native CJK theme typefaces map to matched Noto fonts", () => {
  // The names `applyThemeFontLang` writes into the empty `<a:ea>` slot are the
  // native typeface names from `theme1.xml`, not the romanized ones. Each must
  // resolve to the matching Noto family so measurement and rendering agree, as
  // Japanese already does. (folio bundles no CJK webfonts; the Noto family is a
  // CSS fallback token that resolves to the viewer's OS face when present.)
  const nativeCases: [string, string][] = [
    // Simplified Chinese
    ["宋体", "Noto Serif SC"],
    ["黑体", "Noto Sans SC"],
    ["微软雅黑", "Noto Sans SC"],
    ["等线", "Noto Sans SC"],
    ["仿宋", "Noto Serif SC"],
    ["楷体", "Noto Serif SC"],
    // Traditional Chinese
    ["新細明體", "Noto Serif TC"],
    ["細明體", "Noto Serif TC"],
    ["微軟正黑體", "Noto Sans TC"],
    ["標楷體", "Noto Serif TC"],
    // Korean
    ["맑은 고딕", "Noto Sans KR"],
    ["굴림", "Noto Sans KR"],
    ["돋움", "Noto Sans KR"],
    ["바탕", "Noto Serif KR"],
    ["궁서", "Noto Serif KR"],
    // Japanese native (full-width) — Phase 2 theme path writes these.
    ["ＭＳ 明朝", "Noto Serif JP"],
    ["ＭＳ ゴシック", "Noto Sans JP"],
  ];

  for (const [name, font] of nativeCases) {
    test(`${name} → ${font}`, () => {
      const resolved = resolveFontFamily(name);
      expect(resolved.googleFont).toBe(font);
      expect(resolved.hasGoogleEquivalent).toBe(true);
      expect(getGoogleFontEquivalent(name)).toBe(font);
    });
  }
});

describe("fontResolver — Google Fonts compatibility toggle", () => {
  test("suppresses Google font equivalents without changing local fallback stacks", () => {
    const before = getGoogleFontsEnabled();
    setGoogleFontsEnabled(false);

    try {
      const resolved = resolveFontFamily("Calibri");

      expect(getGoogleFontsEnabled()).toBe(false);
      expect(resolved.googleFont).toBeNull();
      expect(resolved.hasGoogleEquivalent).toBe(false);
      expect(resolved.cssFallback).toContain("Carlito");
      expect(getGoogleFontEquivalent("Calibri")).toBeNull();
      expect(hasGoogleFontEquivalent("Calibri")).toBe(false);
      expect(getGoogleFontsToLoad(["Calibri", "Cambria"])).toEqual([]);
    } finally {
      setGoogleFontsEnabled(before);
    }
  });
});

describe("fontResolver — romanized CJK aliases resolve to the native entry", () => {
  // Word writes the romanized name in run `rFonts`; it must land on the same
  // Noto family + serif/sans category as the native theme name.
  const aliasCases: [string, string, "serif" | "sans-serif"][] = [
    ["SimSun", "Noto Serif SC", "serif"],
    ["Microsoft YaHei", "Noto Sans SC", "sans-serif"],
    ["DengXian", "Noto Sans SC", "sans-serif"],
    ["PMingLiU", "Noto Serif TC", "serif"],
    ["Microsoft JhengHei", "Noto Sans TC", "sans-serif"],
    ["Batang", "Noto Serif KR", "serif"],
    ["Malgun Gothic", "Noto Sans KR", "sans-serif"],
    ["MS Mincho", "Noto Serif JP", "serif"],
    ["Meiryo", "Noto Sans JP", "sans-serif"],
    ["Yu Mincho", "Noto Serif JP", "serif"],
  ];

  for (const [name, font, category] of aliasCases) {
    test(`${name} → ${font} (${category})`, () => {
      const resolved = resolveFontFamily(name);
      expect(resolved.googleFont).toBe(font);
      expect(resolved.hasGoogleEquivalent).toBe(true);
      // A serif Noto family means the run will render with a serif fallback,
      // which is what the serif/sans split must preserve.
      const isSerif = /Noto Serif/u.test(resolved.googleFont ?? "");
      expect(isSerif).toBe(category === "serif");
      expect(hasGoogleFontEquivalent(name)).toBe(true);
    });
  }
});

describe("fontResolver — unmapped native serif faces stay serif", () => {
  // `detectFontCategory` must keep a 明朝/明體/宋 face serif even with no direct
  // mapping, so the generic fallback tail is `serif`, not `sans-serif`.
  for (const name of ["源ノ明朝", "未知明體", "某宋体变体"]) {
    test(`${name} falls back to a serif stack`, () => {
      const resolved = resolveFontFamily(name);
      expect(resolved.cssFallback.endsWith("serif")).toBe(true);
      expect(resolved.cssFallback.includes("sans-serif")).toBe(false);
    });
  }
});

describe("fontResolver — aliased families keep the authored name first", () => {
  // Meiryo / Yu Gothic / Yu Mincho alias to the MS Gothic/Mincho mapping, whose
  // stack does not name them — so the authored family must be prepended for the
  // viewer's own copy to win.
  for (const name of ["Meiryo", "Yu Gothic", "Yu Mincho", "メイリオ"]) {
    test(`${name} is the first fallback`, () => {
      const resolved = resolveFontFamily(name);
      // First family in the stack, with the optional CSS quotes stripped.
      const first = resolved.cssFallback.split(", ")[0]?.replace(/^"|"$/gu, "");
      expect(first).toBe(name);
    });
  }

  test("does not duplicate a name the target stack already lists (PMingLiU)", () => {
    const { cssFallback } = resolveFontFamily("PMingLiU");
    expect(cssFallback.match(/PMingLiU/giu)?.length).toBe(1);
  });
});

describe("fontResolver — untrusted family names are quoted safely", () => {
  test("escapes backslashes and quotes inside quoted CSS family names", () => {
    const { cssFallback } = resolveFontFamily('Evil\\"};x:y');

    expect(cssFallback.startsWith('"')).toBe(true);
    expect(cssFallback).toContain("\\\\");
    expect(cssFallback).toContain('\\"');
    expect(/[^\\]"\};x/u.test(cssFallback)).toBe(false);
  });

  test("hex-escapes CSS newline characters inside quoted family names", () => {
    const { cssFallback } = resolveFontFamily("a\nb\rc\fd");

    expect(cssFallback).not.toMatch(/[\n\r\f]/u);
    expect(cssFallback).toContain("\\a ");
    expect(cssFallback).toContain("\\d ");
    expect(cssFallback).toContain("\\c ");
  });
});
