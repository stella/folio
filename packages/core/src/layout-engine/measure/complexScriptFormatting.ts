type ComplexScriptFormattingSource = {
  complexScriptFontFamily?: string;
  complexScriptAlternateFontFamily?: string;
  complexScriptFontSize?: number;
  complexScriptBold?: boolean;
  complexScriptItalic?: boolean;
};

export type ResolvedComplexScriptFormatting = {
  fontFamily?: string;
  alternateFontFamily?: string;
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
};

/** Resolve Word's independent complex-script slot into ordinary paint fields. */
export const resolveComplexScriptFormatting = (
  source: ComplexScriptFormattingSource,
): ResolvedComplexScriptFormatting => ({
  ...(source.complexScriptFontFamily !== undefined
    ? { fontFamily: source.complexScriptFontFamily }
    : {}),
  ...(source.complexScriptAlternateFontFamily !== undefined
    ? { alternateFontFamily: source.complexScriptAlternateFontFamily }
    : {}),
  ...(source.complexScriptFontSize !== undefined ? { fontSize: source.complexScriptFontSize } : {}),
  ...(source.complexScriptBold !== undefined ? { bold: source.complexScriptBold } : {}),
  ...(source.complexScriptItalic !== undefined ? { italic: source.complexScriptItalic } : {}),
});

export const applyComplexScriptFormatting = <T extends { alternateFontFamily?: string }>(
  base: T,
  source: ComplexScriptFormattingSource,
) => {
  const result = { ...base, ...resolveComplexScriptFormatting(source) };
  if (
    source.complexScriptFontFamily !== undefined &&
    source.complexScriptAlternateFontFamily === undefined
  ) {
    delete result.alternateFontFamily;
  }
  return result;
};

export const hasComplexScriptFormatting = (source: ComplexScriptFormattingSource): boolean =>
  source.complexScriptFontFamily !== undefined ||
  source.complexScriptAlternateFontFamily !== undefined ||
  source.complexScriptFontSize !== undefined ||
  source.complexScriptBold !== undefined ||
  source.complexScriptItalic !== undefined;
