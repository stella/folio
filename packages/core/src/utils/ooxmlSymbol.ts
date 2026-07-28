const OOXML_SYMBOL_CHARACTER_PATTERN = /^[\dA-Fa-f]{4}$/u;

export const isOoxmlSymbolCharacter = (value: string): boolean =>
  OOXML_SYMBOL_CHARACTER_PATTERN.test(value);

export const decodeOoxmlSymbolCharacter = (value: string): string | null =>
  isOoxmlSymbolCharacter(value) ? String.fromCodePoint(Number.parseInt(value, 16)) : null;
