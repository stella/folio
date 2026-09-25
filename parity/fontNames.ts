/** Font-name comparisons shared by the renderer font checks. */

/** PostScript-name decorations a PDF may append to a faithfully-rendered
 * font, normalized to lowercase alphanumerics ("ArialMT" satisfies "Arial",
 * "Calibri-BoldItalic" satisfies "Calibri"). Anything else after the
 * requested family name means the reference renderer used a DIFFERENT family (e.g. a
 * request for "Inter" answered by "Interstate-Bold": remainder "state-bold"
 * is no style suffix, so it counts as a substitution). */
const PDF_FONT_STYLE_SUFFIXES = new Set([
  "",
  "mt",
  "ps",
  "psmt",
  "regular",
  "roman",
  "bold",
  "italic",
  "ital",
  "oblique",
  "bolditalic",
  "boldital",
  "boldoblique",
  "light",
  "medium",
  "semibold",
  "black",
  "boldmt",
  "italicmt",
  "bolditalicmt",
  "psboldmt",
  "psboldital",
  "psitalicmt",
  "psit",
  "psital",
  "psbolditalicmt",
  "psboldital",
]);

export const normalizeFontName = (name: string): string =>
  name
    .replace(/^[A-Z]{6}\+/, "") // PDF subset prefix, e.g. "ABCDEF+Calibri"
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");

export const observedSatisfiesRequested = (observed: string, requested: string): boolean => {
  if (!observed.startsWith(requested)) return false;
  return PDF_FONT_STYLE_SUFFIXES.has(observed.slice(requested.length));
};

/** Whether two renderer-reported names identify the same font family. PDF
 * PostScript names commonly carry style suffixes (`ArialMT`,
 * `Calibri-Bold`), while CSS reports the undecorated family. */
export const fontFamiliesMatch = (leftRaw: string, rightRaw: string): boolean => {
  const left = normalizeFontName(leftRaw);
  const right = normalizeFontName(rightRaw);
  if (left.length === 0 || right.length === 0) return false;
  return observedSatisfiesRequested(left, right) || observedSatisfiesRequested(right, left);
};
