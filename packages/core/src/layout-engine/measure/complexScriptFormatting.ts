type ComplexScriptFormattingSource = {
  complexScriptFontFamily?: string;
  complexScriptFontSize?: number;
  complexScriptBold?: boolean;
  complexScriptItalic?: boolean;
};

export type ResolvedComplexScriptFormatting = {
  fontFamily?: string;
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
  ...(source.complexScriptFontSize !== undefined ? { fontSize: source.complexScriptFontSize } : {}),
  ...(source.complexScriptBold !== undefined ? { bold: source.complexScriptBold } : {}),
  ...(source.complexScriptItalic !== undefined ? { italic: source.complexScriptItalic } : {}),
});

export const hasComplexScriptFormatting = (source: ComplexScriptFormattingSource): boolean =>
  source.complexScriptFontFamily !== undefined ||
  source.complexScriptFontSize !== undefined ||
  source.complexScriptBold !== undefined ||
  source.complexScriptItalic !== undefined;
