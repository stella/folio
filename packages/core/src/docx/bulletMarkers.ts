const SYMBOL_BULLET_MAP: Record<number, string> = {
  0x00_b7: "\u2022",
  0x00_6f: "\u25cb",
  0x00_a7: "\u25a0",
  0x00_fc: "\u2713",
  0x00_6e: "\u25a0",
  0x00_71: "\u25cb",
  0x00_75: "\u25c6",
  0x00_76: "\u2756",
  0x00_a8: "\u2713",
  0x00_fb: "\u2713",
  0x00_fe: "\u2713",
  0xf0_b7: "\u2022",
  0xf0_6e: "\u25a0",
  0xf0_6f: "\u25cb",
  0xf0_a7: "\u25a0",
  0xf0_fc: "\u2713",
  0x20_22: "\u2022",
  0x25_cf: "\u25cf",
  0x25_cb: "\u25cb",
  0x25_a0: "\u25a0",
  0x25_a1: "\u25a1",
  0x25_c6: "\u25c6",
  0x25_c7: "\u25c7",
  0x20_13: "\u2013",
  0x20_14: "\u2014",
  0x00_3e: ">",
  0x00_2d: "-",
};

/**
 * Fonts whose single-byte codes name pictographs rather than Latin letters.
 * A `w:lvlText` character below U+0100 only means a symbol glyph when the
 * numbering level's `w:rFonts` selects one of these faces.
 */
const SYMBOL_ENCODED_FONTS: ReadonlySet<string> = new Set([
  "symbol",
  "wingdings",
  "wingdings 2",
  "wingdings 3",
  "webdings",
]);

type MarkerFontFamily = { ascii?: string | undefined; hAnsi?: string | undefined };

/** The face a bullet level's `w:rFonts` names for its (Latin-range) marker character. */
export const bulletMarkerFontName = (
  formatting: { fontFamily?: MarkerFontFamily | undefined } | null | undefined,
): string | undefined => formatting?.fontFamily?.ascii ?? formatting?.fontFamily?.hAnsi;

/** A printable Latin-1 character in a named, non-symbol face paints as itself. */
const isTextFontCharacter = (charCode: number, fontName: string | undefined): boolean =>
  fontName !== undefined &&
  ((charCode >= 0x20 && charCode < 0x7f) || (charCode >= 0xa0 && charCode < 0x01_00)) &&
  !SYMBOL_ENCODED_FONTS.has(fontName.trim().toLowerCase());

/**
 * Map a bullet level's `w:lvlText` to the character painted for it.
 *
 * `fontName` is the level's own `w:rFonts` face. When it is an ordinary text
 * font, a Latin-range character is the letter itself (the common `o` bullet
 * in a monospace face); without a named face the character is read as a
 * symbol-font code.
 */
export const convertBulletToUnicode = (bulletChar: string, fontName?: string): string => {
  if (!bulletChar || bulletChar.trim() === "") {
    return "\u2022";
  }

  const charCode = bulletChar.codePointAt(0);
  if (charCode === undefined) {
    return "\u2022";
  }

  if (isTextFontCharacter(charCode, fontName)) {
    return bulletChar;
  }

  const mapped = SYMBOL_BULLET_MAP[charCode];
  if (mapped !== undefined) {
    return mapped;
  }

  if (charCode >= 0xe0_00 && charCode <= 0xf8_ff) {
    return "\u2022";
  }

  if (charCode < 32 || (charCode >= 127 && charCode < 160)) {
    return "\u2022";
  }

  return bulletChar;
};
