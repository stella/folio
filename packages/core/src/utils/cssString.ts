/**
 * A CSS `<string>` token, quoted and escaped. Callers emit authored text
 * (a font family from a document) into a stylesheet or an inline style, so
 * the backslash is escaped first, then the quote, and line terminators
 * become hex escapes, which CSS strings require.
 */
export const cssString = (value: string): string =>
  `"${value
    .replaceAll("\\", "\\\\")
    .replaceAll('"', '\\"')
    .replaceAll("\n", "\\A ")
    .replaceAll("\r", "\\D ")
    .replaceAll("\f", "\\C ")}"`;
