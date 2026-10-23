import {
  isOoxmlSymbolFont,
  lookupOoxmlSymbolUnicodeTwin,
} from "../utils/ooxmlSymbol";

/**
 * Fonts whose single-byte codes name pictographs rather than Latin letters.
 * A `w:lvlText` character below U+0100 only means a symbol glyph when the
 * numbering level's `w:rFonts` selects one of these faces.
 */
type MarkerFontFamily = {
  ascii?: string | undefined;
  hAnsi?: string | undefined;
};

/** The face a bullet level's `w:rFonts` names for its (Latin-range) marker character. */
export const bulletMarkerFontName = (
  formatting: { fontFamily?: MarkerFontFamily | undefined } | null | undefined,
): string | undefined =>
  formatting?.fontFamily?.ascii ?? formatting?.fontFamily?.hAnsi;

/** A printable Latin-1 character in a named, non-symbol face paints as itself. */
const isTextFontCharacter = (
  charCode: number,
  fontName: string | undefined,
): boolean =>
  fontName !== undefined &&
  ((charCode >= 0x20 && charCode < 0x7f) ||
    (charCode >= 0xa0 && charCode < 0x01_00)) &&
  !isOoxmlSymbolFont(fontName);

/**
 * Map a bullet level's `w:lvlText` to the character painted for it.
 *
 * `fontName` is the level's own `w:rFonts` face. When it is an ordinary text
 * font, a Latin-range character is the letter itself (the common `o` bullet
 * in a monospace face); without a named face the character is read as a
 * symbol-font code.
 */
export const convertBulletToUnicode = (
  bulletChar: string,
  fontName?: string,
): string => {
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

  const mapped = lookupOoxmlSymbolUnicodeTwin(charCode);
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
