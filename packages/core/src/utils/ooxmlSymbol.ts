import { isOoxmlSymbolCharacter } from "@stll/docx-core/model";

const SYMBOL_BULLET_MAP: ReadonlyMap<number, string> = new Map([
  [0x00_b7, "\u2022"],
  [0x00_6f, "\u25cb"],
  [0x00_a7, "\u25a0"],
  [0x00_fc, "\u2713"],
  [0x00_6e, "\u25a0"],
  [0x00_71, "\u25cb"],
  [0x00_75, "\u25c6"],
  [0x00_76, "\u2756"],
  [0x00_a8, "\u2713"],
  [0x00_fb, "\u2713"],
  [0x00_fe, "\u2713"],
  [0xf0_b7, "\u2022"],
  [0xf0_6e, "\u25a0"],
  [0xf0_6f, "\u25cb"],
  [0xf0_a7, "\u25a0"],
  [0xf0_fc, "\u2713"],
  [0x20_22, "\u2022"],
  [0x25_cf, "\u25cf"],
  [0x25_cb, "\u25cb"],
  [0x25_a0, "\u25a0"],
  [0x25_a1, "\u25a1"],
  [0x25_c6, "\u25c6"],
  [0x25_c7, "\u25c7"],
  [0x20_13, "\u2013"],
  [0x20_14, "\u2014"],
  [0x00_3e, ">"],
  [0x00_2d, "-"],
]);

const SYMBOL_FONT_NAMES: ReadonlySet<string> = new Set([
  "symbol",
  "wingdings",
  "wingdings 2",
  "wingdings 3",
  "webdings",
]);

const SYMBOL_DISPLAY_TWIN_CODES: ReadonlyMap<
  string,
  ReadonlySet<number>
> = new Map([
  ["symbol", new Set([0xf0_b7])],
  ["wingdings", new Set([0xf0_6e, 0xf0_6f, 0xf0_a7, 0xf0_fc])],
]);

export const isOoxmlSymbolFont = (font: string): boolean =>
  SYMBOL_FONT_NAMES.has(font.trim().toLowerCase());

export const lookupOoxmlSymbolUnicodeTwin = (
  codePoint: number,
): string | undefined => SYMBOL_BULLET_MAP.get(codePoint);

export const decodeOoxmlSymbolCharacter = (value: string): string | null =>
  isOoxmlSymbolCharacter(value)
    ? String.fromCodePoint(Number.parseInt(value, 16))
    : null;

export const displayOoxmlSymbolCharacter = (
  font: string,
  value: string,
): string | null => {
  if (!isOoxmlSymbolCharacter(value)) return null;
  const codePoint = Number.parseInt(value, 16);
  const displayTwinCodes = SYMBOL_DISPLAY_TWIN_CODES.get(
    font.trim().toLowerCase(),
  );
  if (!displayTwinCodes?.has(codePoint)) {
    return decodeOoxmlSymbolCharacter(value);
  }
  return (
    lookupOoxmlSymbolUnicodeTwin(codePoint) ?? decodeOoxmlSymbolCharacter(value)
  );
};
