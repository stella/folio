import { isOoxmlSymbolCharacter } from "@stll/docx-core/model";

export const decodeOoxmlSymbolCharacter = (value: string): string | null =>
  isOoxmlSymbolCharacter(value) ? String.fromCodePoint(Number.parseInt(value, 16)) : null;
