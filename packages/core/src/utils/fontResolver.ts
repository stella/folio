/**
 * Font Family Resolver
 *
 * Maps DOCX font names to Google Fonts equivalents with proper CSS fallback stacks.
 *
 * DOCX files often use fonts that aren't freely available (Calibri, Cambria, etc.)
 * This module provides mappings to metrically-compatible Google Fonts alternatives.
 */

import type { ThemeFontScheme, ThemeFont } from "../types";

/**
 * Result of resolving a font family
 */
export type ResolvedFont = {
  /** Google Font name to load (null if no mapping available) */
  googleFont: string | null;
  /** CSS font-family value with proper fallback stack */
  cssFallback: string;
  /** Original font name from the DOCX */
  originalFont: string;
  /** Whether this font has a Google Fonts equivalent */
  hasGoogleEquivalent: boolean;
  /** Single-line height ratio (single-line height ÷ font size). See `FontLineHeight`. */
  singleLineRatio: number;
};

/**
 * Font category for fallback selection
 */
type FontCategory = "sans-serif" | "serif" | "monospace" | "cursive" | "fantasy" | "system-ui";

/**
 * Font mapping entry
 */
type FontMapping = {
  googleFont: string;
  category: FontCategory;
  fallbackStack: string[];
  /** Single-line height ratio (single-line height ÷ font size). See `FontLineHeight`. */
  singleLineRatio: number;
};

/**
 * Default single-line ratio for unmapped fonts.
 * Middle of the common range (1.07–1.27) for standard DOCX fonts.
 */
export const DEFAULT_SINGLE_LINE_RATIO = 1.15;

/**
 * Evidence backing a font's single-line height ratio.
 *
 * - `"hhea"`: the ratio is DERIVED from the font's own `hhea` table metrics
 *   (never hand-written). This is the verified, measured representation —
 *   see `singleLineRatioOf` for the formula and its provenance.
 * - `"legacy"`: an unverified hand-transcribed ratio, kept only for fonts we
 *   have not yet re-measured against real Word output. `note` records why.
 *   Do not add new "legacy" entries without a reason; prefer measuring the
 *   font's `hhea` table instead.
 * - `"measured"`: a ratio measured directly against real Word rendering when
 *   the `hhea` formula does not apply (East-Asian faces: Word derives their
 *   line height from font linking, not the declared font's own `hhea` table).
 *   `note` records the measurement context.
 *
 * Discriminated union so a future edit cannot silently drop the line gap by
 * hand-writing a decimal in its place — every verified font's ratio must
 * trace back to its raw metric fields.
 */
type FontLineHeight =
  | {
      source: "hhea";
      hheaAscent: number;
      hheaDescent: number;
      hheaLineGap: number;
      unitsPerEm: number;
    }
  | { source: "legacy"; ratio: number; note: string }
  | { source: "measured"; ratio: number; note: string };

/**
 * Single-line height ratio (single-line height ÷ font size) for a
 * `FontLineHeight`. This is the ONLY place the derivation formula lives.
 *
 * For `"hhea"` sources: `(hheaAscent + |hheaDescent| + hheaLineGap) / unitsPerEm`.
 * This matches Word's rendered single-line pitch (11pt, single spacing) for
 * every font measured against real Word output so far (16 fonts, no
 * exceptions) — Word does NOT drop the font's line gap. Earlier revisions of
 * this table hand-transcribed ratios from the OS/2 table and omitted the line
 * gap for most fonts, which undershot Word's rendered line height for several
 * of them (confirmed against real Word for cambria, trebuchet ms, palatino
 * linotype, book antiqua, century gothic, consolas, and lucida console; arial
 * and times new roman were fixed in a prior revision). Future readers: to
 * correct or add a font, edit its `hhea*`/`unitsPerEm` metric fields, never the
 * resulting ratio.
 *
 * Most CJK fonts are intentionally left on `DEFAULT_SINGLE_LINE_RATIO`: Word's
 * East-Asian line height is not the run font's hhea ratio (it derives from the
 * paragraph's `w:eastAsia` slot and East-Asian grid layout, not the ascii
 * font), so a single per-font constant cannot capture it correctly. The
 * Japanese Mincho/Gothic entries carry a `"measured"` ratio instead — see
 * `JP_MEASURED_LINE_HEIGHT`.
 */
const singleLineRatioOf = (lineHeight: FontLineHeight): number =>
  lineHeight.source === "hhea"
    ? (lineHeight.hheaAscent - lineHeight.hheaDescent + lineHeight.hheaLineGap) /
      lineHeight.unitsPerEm
    : lineHeight.ratio;

/**
 * Word's East-Asian single-line height for the Japanese Mincho/Gothic faces,
 * measured against real Word output on a NON-grid Japanese document
 * (10.5pt body renders at 13.68pt line pitch → 13.68 / 10.5 ≈ 1.303). Word
 * font-links CJK glyphs to the OS default East-Asian face and takes THAT
 * face's height, so the declared font's own `hhea` table is not the source of
 * truth here; a section with a `w:docGrid` line grid would override this value
 * entirely (out of scope). The ratio is approximate for the CJK long tail
 * (e.g. Yu Mincho renders taller, ≈1.60, when actually installed); it matches
 * the common MS Mincho/Gothic default.
 */
const JP_MEASURED_LINE_HEIGHT: FontLineHeight = {
  source: "measured",
  ratio: 1.303,
  note: "Measured against real Word: 10.5pt Japanese body, non-grid section, renders at 13.68pt line pitch.",
};

/**
 * Mapping of common DOCX fonts to Google Fonts equivalents
 *
 * These are metrically compatible fonts that preserve document layout.
 * See: https://wiki.archlinux.org/title/Metric-compatible_fonts
 *
 * `singleLineRatio` values are computed by `singleLineRatioOf` from each
 * font's `FontLineHeight`, never hand-written. For "hhea" entries the raw
 * `hheaAscent`/`hheaDescent`/`hheaLineGap`/`unitsPerEm` fields below are read
 * directly from the real font files' `hhea` table; the ratio is whatever
 * falls out of the formula in `singleLineRatioOf`. "legacy" entries (e.g.
 * garamond, lucida sans, lucida console) are unverified hand-transcribed
 * ratios carried over unchanged, pending measurement.
 */
const FONT_MAPPINGS: Record<string, FontMapping> = {
  // Microsoft Office fonts -> Google equivalents (via Croscore)
  calibri: {
    googleFont: "Carlito",
    category: "sans-serif",
    fallbackStack: ["Calibri", "Carlito", "Arial", "Helvetica", "sans-serif"],
    singleLineRatio: singleLineRatioOf({
      source: "hhea",
      hheaAscent: 1950,
      hheaDescent: -550,
      hheaLineGap: 0,
      unitsPerEm: 2048,
    }), // 1.2207
  },
  cambria: {
    googleFont: "Caladea",
    category: "serif",
    fallbackStack: ["Cambria", "Caladea", "Georgia", "serif"],
    singleLineRatio: singleLineRatioOf({
      source: "hhea",
      hheaAscent: 1946,
      hheaDescent: -455,
      hheaLineGap: 0,
      unitsPerEm: 2048,
    }), // 1.1724 (was hand-transcribed 1.2676 — wrong by 8%)
  },
  arial: {
    googleFont: "Arimo",
    category: "sans-serif",
    fallbackStack: ["Arial", "Arimo", "Helvetica", "sans-serif"],
    singleLineRatio: singleLineRatioOf({
      source: "hhea",
      hheaAscent: 1854,
      hheaDescent: -434,
      hheaLineGap: 67,
      unitsPerEm: 2048,
    }), // 1.1499
  },
  "times new roman": {
    googleFont: "Tinos",
    category: "serif",
    fallbackStack: ["Times New Roman", "Tinos", "Times", "serif"],
    singleLineRatio: singleLineRatioOf({
      source: "hhea",
      hheaAscent: 1825,
      hheaDescent: -443,
      hheaLineGap: 87,
      unitsPerEm: 2048,
    }), // 1.1499
  },
  "cg times": {
    googleFont: "Tinos",
    category: "serif",
    fallbackStack: ["Times New Roman", "Tinos", "Times", "serif"],
    singleLineRatio: singleLineRatioOf({
      source: "hhea",
      hheaAscent: 1825,
      hheaDescent: -443,
      hheaLineGap: 87,
      unitsPerEm: 2048,
    }), // Metric-compatible Times New Roman substitute used by Word/LibreOffice.
  },
  "courier new": {
    googleFont: "Cousine",
    category: "monospace",
    fallbackStack: ["Courier New", "Cousine", "Courier", "monospace"],
    singleLineRatio: singleLineRatioOf({
      source: "hhea",
      hheaAscent: 1705,
      hheaDescent: -615,
      hheaLineGap: 0,
      unitsPerEm: 2048,
    }), // 1.1328
  },

  // Additional common fonts
  georgia: {
    googleFont: "Tinos", // Similar but not perfect match
    category: "serif",
    fallbackStack: ["Georgia", "Tinos", "Times New Roman", "serif"],
    singleLineRatio: singleLineRatioOf({
      source: "hhea",
      hheaAscent: 1878,
      hheaDescent: -449,
      hheaLineGap: 0,
      unitsPerEm: 2048,
    }), // 1.1362
  },
  verdana: {
    googleFont: "Open Sans", // Similar sans-serif
    category: "sans-serif",
    fallbackStack: ["Verdana", "Open Sans", "Arial", "sans-serif"],
    singleLineRatio: singleLineRatioOf({
      source: "hhea",
      hheaAscent: 2059,
      hheaDescent: -430,
      hheaLineGap: 0,
      unitsPerEm: 2048,
    }), // 1.2153
  },
  tahoma: {
    googleFont: "Open Sans",
    category: "sans-serif",
    fallbackStack: ["Tahoma", "Open Sans", "Arial", "sans-serif"],
    singleLineRatio: singleLineRatioOf({
      source: "hhea",
      hheaAscent: 2049,
      hheaDescent: -423,
      hheaLineGap: 0,
      unitsPerEm: 2048,
    }), // 1.2070 (was hand-transcribed 1.2075 — negligible, now derived)
  },
  "trebuchet ms": {
    googleFont: "Fira Sans",
    category: "sans-serif",
    fallbackStack: ["Trebuchet MS", "Fira Sans", "Arial", "sans-serif"],
    singleLineRatio: singleLineRatioOf({
      source: "hhea",
      hheaAscent: 1923,
      hheaDescent: -455,
      hheaLineGap: 0,
      unitsPerEm: 2048,
    }), // 1.1611 (was hand-transcribed 1.1431 — wrong)
  },
  "comic sans ms": {
    googleFont: "Comic Neue",
    category: "cursive",
    fallbackStack: ["Comic Sans MS", "Comic Neue", "cursive"],
    singleLineRatio: singleLineRatioOf({
      source: "hhea",
      hheaAscent: 2257,
      hheaDescent: -597,
      hheaLineGap: 0,
      unitsPerEm: 2048,
    }), // 1.3936
  },
  impact: {
    googleFont: "Anton",
    category: "sans-serif",
    fallbackStack: ["Impact", "Anton", "Arial Black", "sans-serif"],
    singleLineRatio: singleLineRatioOf({
      source: "hhea",
      hheaAscent: 2066,
      hheaDescent: -432,
      hheaLineGap: 0,
      unitsPerEm: 2048,
    }), // 1.2197
  },
  "palatino linotype": {
    googleFont: "EB Garamond",
    category: "serif",
    fallbackStack: ["Palatino Linotype", "EB Garamond", "Palatino", "Georgia", "serif"],
    singleLineRatio: singleLineRatioOf({
      source: "hhea",
      hheaAscent: 2150,
      hheaDescent: -613,
      hheaLineGap: 0,
      unitsPerEm: 2048,
    }), // 1.3491 (was hand-transcribed 1.0259 — WRONG by 31%)
  },
  "book antiqua": {
    googleFont: "EB Garamond",
    category: "serif",
    fallbackStack: ["Book Antiqua", "EB Garamond", "Palatino", "Georgia", "serif"],
    singleLineRatio: singleLineRatioOf({
      source: "hhea",
      hheaAscent: 1891,
      hheaDescent: -578,
      hheaLineGap: 0,
      unitsPerEm: 2048,
    }), // 1.2056 (was hand-transcribed 1.0259 — wrong by 17%)
  },
  garamond: {
    googleFont: "EB Garamond",
    category: "serif",
    fallbackStack: ["Garamond", "EB Garamond", "Georgia", "serif"],
    singleLineRatio: singleLineRatioOf({
      source: "legacy",
      ratio: 1.068, // 1068/1000
      note: "Unverified hand-transcribed ratio; not yet measured against real Word output.",
    }),
  },
  "century gothic": {
    googleFont: "Questrial",
    category: "sans-serif",
    fallbackStack: ["Century Gothic", "Questrial", "Arial", "sans-serif"],
    singleLineRatio: singleLineRatioOf({
      source: "hhea",
      hheaAscent: 2060,
      hheaDescent: -451,
      hheaLineGap: 0,
      unitsPerEm: 2048,
    }), // 1.2261 (was hand-transcribed 1.1611 — wrong by 6%)
  },
  "lucida sans": {
    googleFont: "Open Sans",
    category: "sans-serif",
    fallbackStack: ["Lucida Sans", "Open Sans", "Arial", "sans-serif"],
    singleLineRatio: singleLineRatioOf({
      source: "legacy",
      ratio: 1.1655, // 2387/2048
      note: "Unverified hand-transcribed ratio; not yet measured against real Word output.",
    }),
  },
  "lucida console": {
    googleFont: "Cousine",
    category: "monospace",
    fallbackStack: ["Lucida Console", "Cousine", "Courier New", "monospace"],
    singleLineRatio: singleLineRatioOf({
      source: "hhea",
      hheaAscent: 1616,
      hheaDescent: -432,
      hheaLineGap: 0,
      unitsPerEm: 2048,
    }), // 1.0000 (was hand-transcribed 1.1387 — wrong; Word renders single-spaced Lucida Console at 1.0×)
  },
  consolas: {
    googleFont: "Inconsolata",
    category: "monospace",
    fallbackStack: ["Consolas", "Inconsolata", "Cousine", "Courier New", "monospace"],
    singleLineRatio: singleLineRatioOf({
      source: "hhea",
      hheaAscent: 1521,
      hheaDescent: -527,
      hheaLineGap: 350,
      unitsPerEm: 2048,
    }), // 1.1709 (was hand-transcribed 1.1626 — wrong)
  },

  // CJK fonts
  "ms mincho": {
    googleFont: "Noto Serif JP",
    category: "serif",
    fallbackStack: ["MS Mincho", "Noto Serif JP", "serif"],
    singleLineRatio: singleLineRatioOf(JP_MEASURED_LINE_HEIGHT),
  },
  // Native Japanese typeface names as they appear in `theme1.xml`'s
  // `<a:font script="Jpan">` entries (full-width "ＭＳ"). Office stores these
  // rather than the romanized "MS Mincho"/"MS Gothic" names, so map them too.
  // NOTE: keys are matched against `name.toLowerCase()`, and full-width Latin
  // letters lowercase too (Ｍ→ｍ, Ｓ→ｓ, Ｐ→ｐ) — so these keys are lowercased.
  "ｍｓ 明朝": {
    googleFont: "Noto Serif JP",
    category: "serif",
    fallbackStack: ["MS Mincho", "ＭＳ 明朝", "Noto Serif JP", "serif"],
    singleLineRatio: singleLineRatioOf(JP_MEASURED_LINE_HEIGHT),
  },
  "ｍｓ ｐ明朝": {
    googleFont: "Noto Serif JP",
    category: "serif",
    fallbackStack: ["MS PMincho", "ＭＳ Ｐ明朝", "Noto Serif JP", "serif"],
    singleLineRatio: singleLineRatioOf(JP_MEASURED_LINE_HEIGHT),
  },
  "ms gothic": {
    googleFont: "Noto Sans JP",
    category: "sans-serif",
    fallbackStack: ["MS Gothic", "Noto Sans JP", "sans-serif"],
    singleLineRatio: singleLineRatioOf(JP_MEASURED_LINE_HEIGHT),
  },
  "ｍｓ ゴシック": {
    googleFont: "Noto Sans JP",
    category: "sans-serif",
    fallbackStack: ["MS Gothic", "ＭＳ ゴシック", "Noto Sans JP", "sans-serif"],
    singleLineRatio: singleLineRatioOf(JP_MEASURED_LINE_HEIGHT),
  },
  "ｍｓ ｐゴシック": {
    googleFont: "Noto Sans JP",
    category: "sans-serif",
    fallbackStack: ["MS PGothic", "ＭＳ Ｐゴシック", "Noto Sans JP", "sans-serif"],
    singleLineRatio: singleLineRatioOf(JP_MEASURED_LINE_HEIGHT),
  },
  simhei: {
    googleFont: "Noto Sans SC",
    category: "sans-serif",
    fallbackStack: ["SimHei", "Noto Sans SC", "sans-serif"],
    singleLineRatio: DEFAULT_SINGLE_LINE_RATIO,
  },
  simsun: {
    googleFont: "Noto Serif SC",
    category: "serif",
    fallbackStack: ["SimSun", "Noto Serif SC", "serif"],
    singleLineRatio: DEFAULT_SINGLE_LINE_RATIO,
  },
  "malgun gothic": {
    googleFont: "Noto Sans KR",
    category: "sans-serif",
    fallbackStack: ["Malgun Gothic", "Noto Sans KR", "sans-serif"],
    singleLineRatio: DEFAULT_SINGLE_LINE_RATIO,
  },
  // Native CJK typeface names as stored in `theme1.xml`'s `<a:font script="…">`
  // entries (port of eigenpal/docx-editor#957). `applyThemeFontLang` resolves
  // empty `<a:ea>` slots to these native names (not the romanized
  // "SimSun"/"Malgun Gothic"), so map them to the same Noto families. folio
  // bundles no CJK webfonts: the Noto family resolves to the viewer's OS face
  // when present, else the generic serif/sans tail. (CJK has no letter case, so
  // `.toLowerCase()` leaves these keys unchanged.)

  // Simplified Chinese
  宋体: {
    googleFont: "Noto Serif SC",
    category: "serif",
    fallbackStack: ["SimSun", "宋体", "Noto Serif SC", "serif"],
    singleLineRatio: DEFAULT_SINGLE_LINE_RATIO,
  },
  黑体: {
    googleFont: "Noto Sans SC",
    category: "sans-serif",
    fallbackStack: ["SimHei", "黑体", "Noto Sans SC", "sans-serif"],
    singleLineRatio: DEFAULT_SINGLE_LINE_RATIO,
  },
  微软雅黑: {
    googleFont: "Noto Sans SC",
    category: "sans-serif",
    fallbackStack: ["Microsoft YaHei", "微软雅黑", "Noto Sans SC", "sans-serif"],
    singleLineRatio: DEFAULT_SINGLE_LINE_RATIO,
  },
  等线: {
    googleFont: "Noto Sans SC",
    category: "sans-serif",
    fallbackStack: ["DengXian", "等线", "Noto Sans SC", "sans-serif"],
    singleLineRatio: DEFAULT_SINGLE_LINE_RATIO,
  },
  仿宋: {
    googleFont: "Noto Serif SC",
    category: "serif",
    fallbackStack: ["FangSong", "仿宋", "Noto Serif SC", "serif"],
    singleLineRatio: DEFAULT_SINGLE_LINE_RATIO,
  },
  楷体: {
    googleFont: "Noto Serif SC",
    category: "serif",
    fallbackStack: ["KaiTi", "楷体", "Noto Serif SC", "serif"],
    singleLineRatio: DEFAULT_SINGLE_LINE_RATIO,
  },

  // Traditional Chinese
  新細明體: {
    googleFont: "Noto Serif TC",
    category: "serif",
    fallbackStack: ["PMingLiU", "新細明體", "Noto Serif TC", "serif"],
    singleLineRatio: DEFAULT_SINGLE_LINE_RATIO,
  },
  細明體: {
    googleFont: "Noto Serif TC",
    category: "serif",
    fallbackStack: ["MingLiU", "細明體", "Noto Serif TC", "serif"],
    singleLineRatio: DEFAULT_SINGLE_LINE_RATIO,
  },
  微軟正黑體: {
    googleFont: "Noto Sans TC",
    category: "sans-serif",
    fallbackStack: ["Microsoft JhengHei", "微軟正黑體", "Noto Sans TC", "sans-serif"],
    singleLineRatio: DEFAULT_SINGLE_LINE_RATIO,
  },
  標楷體: {
    googleFont: "Noto Serif TC",
    category: "serif",
    fallbackStack: ["DFKai-SB", "標楷體", "Noto Serif TC", "serif"],
    singleLineRatio: DEFAULT_SINGLE_LINE_RATIO,
  },

  // Korean
  "맑은 고딕": {
    googleFont: "Noto Sans KR",
    category: "sans-serif",
    fallbackStack: ["Malgun Gothic", "맑은 고딕", "Noto Sans KR", "sans-serif"],
    singleLineRatio: DEFAULT_SINGLE_LINE_RATIO,
  },
  굴림: {
    googleFont: "Noto Sans KR",
    category: "sans-serif",
    fallbackStack: ["Gulim", "굴림", "Noto Sans KR", "sans-serif"],
    singleLineRatio: DEFAULT_SINGLE_LINE_RATIO,
  },
  돋움: {
    googleFont: "Noto Sans KR",
    category: "sans-serif",
    fallbackStack: ["Dotum", "돋움", "Noto Sans KR", "sans-serif"],
    singleLineRatio: DEFAULT_SINGLE_LINE_RATIO,
  },
  바탕: {
    googleFont: "Noto Serif KR",
    category: "serif",
    fallbackStack: ["Batang", "바탕", "Noto Serif KR", "serif"],
    singleLineRatio: DEFAULT_SINGLE_LINE_RATIO,
  },
  궁서: {
    googleFont: "Noto Serif KR",
    category: "serif",
    fallbackStack: ["Gungsuh", "궁서", "Noto Serif KR", "serif"],
    singleLineRatio: DEFAULT_SINGLE_LINE_RATIO,
  },
};

/**
 * Default fallback stacks by category
 */
const DEFAULT_FALLBACKS = {
  "sans-serif": "Arial, Helvetica, sans-serif",
  serif: "Times New Roman, Times, serif",
  monospace: "Courier New, Courier, monospace",
  cursive: "cursive",
  fantasy: "fantasy",
  "system-ui": "system-ui, sans-serif",
} as const satisfies Record<FontCategory, string>;

/**
 * Detect font category from font name
 */
function detectFontCategory(fontName: string): FontCategory {
  const lower = fontName.toLowerCase();

  // Monospace indicators
  if (
    lower.includes("mono") ||
    lower.includes("courier") ||
    lower.includes("consolas") ||
    lower.includes("console") ||
    lower.includes("code") ||
    lower.includes("terminal")
  ) {
    return "monospace";
  }

  // Serif indicators
  if (
    lower.includes("times") ||
    lower.includes("georgia") ||
    lower.includes("garamond") ||
    lower.includes("palatino") ||
    lower.includes("baskerville") ||
    lower.includes("bodoni") ||
    lower.includes("cambria") ||
    lower.includes("mincho") ||
    lower.includes("明朝") ||
    lower.includes("明體") ||
    lower.includes("宋") ||
    lower.includes("ming") ||
    lower.includes("song") ||
    lower.includes("serif")
  ) {
    return "serif";
  }

  // Cursive/script indicators
  if (
    lower.includes("script") ||
    lower.includes("cursive") ||
    lower.includes("comic") ||
    lower.includes("brush") ||
    lower.includes("hand")
  ) {
    return "cursive";
  }

  // Default to sans-serif
  return "sans-serif";
}

/**
 * Romanized CJK font spellings → the native key already present in
 * FONT_MAPPINGS (port of eigenpal/docx-editor#957). Word writes either the
 * native typeface name (the theme path, e.g. `新細明體`) or the romanized name
 * (run `rFonts`, e.g. `PMingLiU`); both must resolve to the same Noto family +
 * CSS stack + category. Aliasing keeps a single set of CJK entries instead of
 * duplicating each under both spellings. (Keys already present directly in
 * FONT_MAPPINGS — `simsun`, `simhei`, `malgun gothic`, `ms mincho`,
 * `ms gothic` — are intentionally omitted.)
 */
const CJK_FONT_ALIASES: Record<string, string> = {
  // Simplified Chinese
  "microsoft yahei": "微软雅黑",
  dengxian: "等线",
  fangsong: "仿宋",
  kaiti: "楷体",
  // Traditional Chinese
  pmingliu: "新細明體",
  mingliu: "細明體",
  "microsoft jhenghei": "微軟正黑體",
  "dfkai-sb": "標楷體",
  // Korean
  gulim: "굴림",
  dotum: "돋움",
  batang: "바탕",
  gungsuh: "궁서",
  // Japanese — Meiryo / Yu share Noto Sans/Serif JP with MS Gothic/Mincho
  meiryo: "ms gothic",
  メイリオ: "ms gothic",
  "yu gothic": "ms gothic",
  游ゴシック: "ms gothic",
  "yu mincho": "ms mincho",
  游明朝: "ms mincho",
};

/**
 * The Noto families every CJK entry in `FONT_MAPPINGS` maps to. A mapping
 * whose `googleFont` is one of these is an East-Asian face; this is the single
 * classification source for `isCjkFont`, kept as an explicit set (not a name
 * heuristic) so a future non-Noto CJK mapping must extend it deliberately.
 */
const CJK_NOTO_FAMILIES: ReadonlySet<string> = new Set([
  "Noto Serif JP",
  "Noto Sans JP",
  "Noto Serif SC",
  "Noto Sans SC",
  "Noto Serif TC",
  "Noto Sans TC",
  "Noto Serif KR",
  "Noto Sans KR",
]);

/**
 * True when `family` is a known East-Asian (CJK) typeface — a direct
 * `FONT_MAPPINGS` CJK entry or a romanized alias of one. Used by the measurer
 * to decide whether a run's `w:eastAsia` font can supply the line height for
 * CJK text, or whether Word would font-link to a default East-Asian face
 * instead (see `CJK_FALLBACK_FONT_FAMILY`). Unmapped families return false:
 * we cannot know their metrics, so the caller falls back.
 */
export function isCjkFont(family: string): boolean {
  const normalizedName = family.trim().toLowerCase();
  const mapping = FONT_MAPPINGS[CJK_FONT_ALIASES[normalizedName] ?? normalizedName];
  return mapping !== undefined && CJK_NOTO_FAMILIES.has(mapping.googleFont);
}

/**
 * Font family whose `singleLineRatio` stands in for Word's default East-Asian
 * face when a CJK-text run declares no usable CJK font (its `w:eastAsia` slot
 * is absent or names a Latin face like "Century"). Word font-links those
 * glyphs to the OS default East-Asian font and takes THAT face's taller line
 * height; MS Mincho carries the measured ratio for it
 * (`JP_MEASURED_LINE_HEIGHT`, ≈1.303). Line-height resolution only — never
 * used for width measurement or painting.
 */
export const CJK_FALLBACK_FONT_FAMILY = "MS Mincho";

/**
 * Resolve a DOCX font name to a Google Font and CSS fallback stack
 *
 * @param docxFontName - The font name from the DOCX file
 * @returns Resolved font information
 */
// The bundled Latin (Croscore) faces carry no Arabic glyphs, so weave a
// bundled Arabic face into every fallback chain just before the generic
// family. Latin text keeps its named font; Arabic runs fall through to Noto.
const ARABIC_FALLBACK = '"Noto Sans Arabic"';
const TRAILING_GENERIC = /,\s*(?<generic>sans-serif|serif|monospace|cursive)\s*$/u;
const withArabicFallback = (stack: string): string =>
  TRAILING_GENERIC.test(stack)
    ? stack.replace(TRAILING_GENERIC, `, ${ARABIC_FALLBACK}, $<generic>`)
    : `${stack}, ${ARABIC_FALLBACK}`;

let googleFontsEnabled = true;

export const setGoogleFontsEnabled = (enabled: boolean): void => {
  googleFontsEnabled = enabled;
};

export const getGoogleFontsEnabled = (): boolean => googleFontsEnabled;

export function resolveFontFamily(docxFontName: string): ResolvedFont {
  const normalizedName = docxFontName.trim().toLowerCase();

  // Direct mapping, or a romanized CJK spelling aliased to its native entry.
  const aliasTarget = CJK_FONT_ALIASES[normalizedName];
  const mapping = FONT_MAPPINGS[aliasTarget ?? normalizedName];

  if (mapping) {
    // When reached via an alias, the authored family may be absent from the
    // target's stack (e.g. Meiryo / Yu Gothic alias to the MS Gothic mapping).
    // Prepend it so the viewer's own copy of the named face is tried first.
    const fallbackStack =
      aliasTarget && !mapping.fallbackStack.some((f) => f.toLowerCase() === normalizedName)
        ? [docxFontName, ...mapping.fallbackStack]
        : mapping.fallbackStack;
    return {
      googleFont: googleFontsEnabled ? mapping.googleFont : null,
      cssFallback: withArabicFallback(fallbackStack.map(quoteFontName).join(", ")),
      originalFont: docxFontName,
      hasGoogleEquivalent: googleFontsEnabled,
      singleLineRatio: mapping.singleLineRatio,
    };
  }

  // No mapping - detect category and create fallback
  const category = detectFontCategory(docxFontName);
  const defaultFallback = DEFAULT_FALLBACKS[category];

  return {
    googleFont: null,
    cssFallback: withArabicFallback(`${quoteFontName(docxFontName)}, ${defaultFallback}`),
    originalFont: docxFontName,
    hasGoogleEquivalent: false,
    singleLineRatio: DEFAULT_SINGLE_LINE_RATIO,
  };
}

const CSS_NEWLINE_ESCAPES: Record<string, string> = {
  "\n": "\\a ",
  "\r": "\\d ",
  "\f": "\\c ",
};

/**
 * Escape a DOCX-supplied font name for a double-quoted CSS string.
 */
function escapeQuotedFontName(fontName: string): string {
  return fontName
    .replace(/["\\]/g, "\\$&")
    .replace(/[\n\r\f]/g, (char) => CSS_NEWLINE_ESCAPES[char] ?? char);
}

/**
 * Quote a font name if it contains spaces or special characters
 */
function quoteFontName(fontName: string): string {
  // If it's a generic family, don't quote
  if (
    ["serif", "sans-serif", "monospace", "cursive", "fantasy", "system-ui"].includes(
      fontName.toLowerCase(),
    )
  ) {
    return fontName;
  }

  // Quote if contains spaces or special characters
  if (/[\s,'"()]/.test(fontName)) {
    return `"${escapeQuotedFontName(fontName)}"`;
  }

  return fontName;
}

/**
 * Resolve a theme font reference to actual font names
 *
 * @param themeRef - Theme font reference (e.g., 'majorAscii', 'minorHAnsi')
 * @param fontScheme - Theme font scheme from the DOCX
 * @returns Resolved font name or null if not found
 */
export function resolveThemeFont(themeRef: string, fontScheme?: ThemeFontScheme): string | null {
  if (!fontScheme) {
    return null;
  }

  // Parse the theme reference
  const isMajor = themeRef.toLowerCase().startsWith("major");
  const themeFont: ThemeFont | undefined = isMajor ? fontScheme.majorFont : fontScheme.minorFont;

  if (!themeFont) {
    return null;
  }

  // Determine which script variant to use
  const ref = themeRef.toLowerCase();

  if (ref.includes("eastasia") || ref.includes("ea")) {
    return themeFont.ea ?? themeFont.latin ?? null;
  }

  if (ref.includes("cs") || ref.includes("bidi")) {
    return themeFont.cs ?? themeFont.latin ?? null;
  }

  // Default to Latin/Western font
  return themeFont.latin ?? null;
}

/**
 * Get all Google Font names needed for a list of DOCX fonts
 *
 * @param docxFonts - Array of font names from the DOCX
 * @returns Array of Google Font names to load
 */
export function getGoogleFontsToLoad(docxFonts: string[]): string[] {
  const googleFonts = new Set<string>();

  for (const font of docxFonts) {
    const resolved = resolveFontFamily(font);
    if (resolved.googleFont) {
      googleFonts.add(resolved.googleFont);
    }
  }

  return Array.from(googleFonts);
}

/**
 * Build a CSS font-family string from an array of font names
 *
 * @param fonts - Array of font names
 * @param category - Fallback category
 * @returns CSS font-family value
 */
export function buildFontFamilyString(
  fonts: string[],
  category: FontCategory = "sans-serif",
): string {
  const quoted = fonts.map(quoteFontName);

  // Ensure we have the generic fallback
  const parts = [...quoted];
  if (!parts.some((p) => p.toLowerCase() === category)) {
    parts.push(category);
  }

  return parts.join(", ");
}

/**
 * Get the Google Font equivalent for a DOCX font (if any)
 *
 * @param docxFontName - The font name from the DOCX file
 * @returns Google Font name or null
 */
export function getGoogleFontEquivalent(docxFontName: string): string | null {
  // Delegate so CJK romanized aliases are honored identically to the CSS path.
  return resolveFontFamily(docxFontName).googleFont;
}

/**
 * Check if a font has a known Google Fonts equivalent
 *
 * @param docxFontName - The font name from the DOCX file
 * @returns true if there's a Google Fonts equivalent
 */
export function hasGoogleFontEquivalent(docxFontName: string): boolean {
  if (!googleFontsEnabled) {
    return false;
  }

  const normalizedName = docxFontName.trim().toLowerCase();
  const aliasTarget = CJK_FONT_ALIASES[normalizedName];
  return (aliasTarget ?? normalizedName) in FONT_MAPPINGS;
}
