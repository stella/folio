/**
 * Script segmentation for per-character font selection.
 *
 * Word resolves a run's font per character across three slots: East-Asian code
 * points use `w:eastAsia`, complex-script code points (Arabic, Hebrew, Indic,
 * South-East Asian) use `w:cs`, and everything else uses `w:ascii`/`w:hAnsi`.
 * folio mirrors this by splitting a run's text into maximal same-script
 * segments that the measurer and the painter both consume, so line wrapping
 * stays in sync with rendering.
 *
 * A three-value discriminator rather than the `isCjk` boolean this started as:
 * the question is "which font slot does this character take?", which already
 * has three answers and may grow again.
 *
 * Ranges are authored with `\u` escapes only — never pasted glyphs (a pasted
 * glyph once silently corrupted an RTL character class here).
 */

/** Which of Word's font slots a code point selects. */
export const SCRIPT_CLASS = {
  /** `w:eastAsia` — CJK ideographs, kana, Hangul. */
  eastAsia: "eastAsia",
  /** `w:cs` — Arabic, Hebrew, Indic, South-East Asian. */
  complex: "complex",
  /** `w:ascii` / `w:hAnsi` — everything else. */
  western: "western",
} as const;

export type ScriptClass = (typeof SCRIPT_CLASS)[keyof typeof SCRIPT_CLASS];

export type ScriptSegment = {
  text: string;
  script: ScriptClass;
};

/**
 * True when a code point belongs to an East-Asian script that Word renders with
 * the `w:eastAsia` font slot: CJK ideographs, kana, Hangul, CJK symbols and
 * punctuation, and fullwidth/halfwidth forms.
 */
export function isCjkCodePoint(cp: number): boolean {
  return (
    (cp >= 0x2e_80 && cp <= 0x2f_ff) || // CJK & Kangxi radicals, ideographic description
    (cp >= 0x30_00 && cp <= 0x33_ff) || // symbols/punctuation, kana, Bopomofo, Hangul compat Jamo, Kanbun, Bopomofo ext, CJK strokes, Katakana phonetic ext, enclosed/compatibility CJK
    (cp >= 0x34_00 && cp <= 0x4d_bf) || // CJK Unified Ideographs Extension A
    (cp >= 0x4e_00 && cp <= 0x9f_ff) || // CJK Unified Ideographs
    (cp >= 0xa9_60 && cp <= 0xa9_7f) || // Hangul Jamo Extended-A
    (cp >= 0xac_00 && cp <= 0xd7_ff) || // Hangul syllables + Jamo Extended-B
    (cp >= 0x11_00 && cp <= 0x11_ff) || // Hangul Jamo
    (cp >= 0xf9_00 && cp <= 0xfa_ff) || // CJK compatibility ideographs
    (cp >= 0xfe_30 && cp <= 0xfe_4f) || // CJK compatibility forms
    (cp >= 0xff_00 && cp <= 0xff_ef) || // Halfwidth and fullwidth forms
    (cp >= 0x2_00_00 && cp <= 0x3_ff_ff) // CJK Unified Ideographs Extension B+ (astral)
  );
}

/**
 * True when the text contains at least one East-Asian code point. Callers use
 * this to skip segmentation entirely on the common all-Latin path.
 */
export function hasCjk(text: string): boolean {
  for (const ch of text) {
    // SAFETY: for...of over a string yields whole code points.
    if (isCjkCodePoint(ch.codePointAt(0)!)) {
      return true;
    }
  }
  return false;
}

/**
 * True when a code point belongs to a script Word renders with the `w:cs`
 * (complex script) font slot: the bidirectional scripts plus the Indic and
 * South-East Asian scripts that need contextual shaping.
 *
 * Checked after CJK, which owns the ranges above U+2E80 that would otherwise
 * overlap the presentation forms below.
 */
export function isComplexScriptCodePoint(cp: number): boolean {
  return (
    (cp >= 0x05_90 && cp <= 0x05_ff) || // Hebrew
    (cp >= 0x06_00 && cp <= 0x06_ff) || // Arabic
    (cp >= 0x07_00 && cp <= 0x07_4f) || // Syriac
    (cp >= 0x07_50 && cp <= 0x07_7f) || // Arabic Supplement
    (cp >= 0x07_80 && cp <= 0x07_bf) || // Thaana
    (cp >= 0x07_c0 && cp <= 0x07_ff) || // NKo
    (cp >= 0x08_00 && cp <= 0x08_3f) || // Samaritan
    (cp >= 0x08_40 && cp <= 0x08_5f) || // Mandaic
    (cp >= 0x08_60 && cp <= 0x08_6f) || // Syriac Supplement
    (cp >= 0x08_70 && cp <= 0x08_9f) || // Arabic Extended-B
    (cp >= 0x08_a0 && cp <= 0x08_ff) || // Arabic Extended-A
    (cp >= 0x09_00 && cp <= 0x0d_7f) || // Devanagari through Malayalam
    (cp >= 0x0d_80 && cp <= 0x0d_ff) || // Sinhala
    (cp >= 0x0e_00 && cp <= 0x0e_7f) || // Thai
    (cp >= 0x0e_80 && cp <= 0x0e_ff) || // Lao
    (cp >= 0x0f_00 && cp <= 0x0f_ff) || // Tibetan
    (cp >= 0x10_00 && cp <= 0x10_9f) || // Myanmar
    (cp >= 0x17_80 && cp <= 0x17_ff) || // Khmer
    (cp >= 0xfb_1d && cp <= 0xfb_4f) || // Hebrew presentation forms
    (cp >= 0xfb_50 && cp <= 0xfd_ff) || // Arabic presentation forms-A
    (cp >= 0xfe_70 && cp <= 0xfe_ff) || // Arabic presentation forms-B
    // Astral blocks. The BMP ranges above stop at U+FEFF, which silently left
    // every astral Arabic-family script selecting the western slot.
    (cp >= 0x10_d0_0 && cp <= 0x10_d3_f) || // Hanifi Rohingya
    (cp >= 0x10_ec_0 && cp <= 0x10_ef_f) || // Arabic Extended-C
    (cp >= 0x10_f3_0 && cp <= 0x10_f6_f) || // Sogdian
    (cp >= 0x10_f7_0 && cp <= 0x10_fa_f) || // Old Uyghur
    (cp >= 0x1e_90_0 && cp <= 0x1e_95_f) || // Adlam
    (cp >= 0x1e_e0_0 && cp <= 0x1e_ef_f) // Arabic Mathematical Alphabetic Symbols
  );
}

/** Which font slot a single code point selects. CJK wins where ranges meet. */
export function scriptClassOf(cp: number): ScriptClass {
  if (isCjkCodePoint(cp)) {
    return SCRIPT_CLASS.eastAsia;
  }
  if (isComplexScriptCodePoint(cp)) {
    return SCRIPT_CLASS.complex;
  }
  return SCRIPT_CLASS.western;
}

/**
 * True when the text contains at least one complex-script code point. Callers
 * use this to skip segmentation entirely on the common all-Latin path.
 */
export function hasComplexScript(text: string): boolean {
  for (const ch of text) {
    // SAFETY: for...of over a string yields whole code points.
    if (isComplexScriptCodePoint(ch.codePointAt(0)!)) {
      return true;
    }
  }
  return false;
}

/**
 * Split text into maximal runs of one script class (CJK vs. non-CJK). Iterates
 * by code point so astral ideographs (surrogate pairs) are never split between
 * fonts. Empty input yields no segments; single-class input yields one.
 */
export function segmentByScript(text: string): ScriptSegment[] {
  const segments: ScriptSegment[] = [];
  let current = "";
  let currentScript: ScriptClass = SCRIPT_CLASS.western;

  for (const ch of text) {
    // SAFETY: for...of over a string yields whole code points.
    const script = scriptClassOf(ch.codePointAt(0)!);
    if (current.length === 0) {
      current = ch;
      currentScript = script;
      continue;
    }
    if (script === currentScript) {
      current += ch;
      continue;
    }
    segments.push({ text: current, script: currentScript });
    current = ch;
    currentScript = script;
  }

  if (current.length > 0) {
    segments.push({ text: current, script: currentScript });
  }

  return segments;
}
