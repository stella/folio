/** Private error sink: retain repeated refusals even when the status text is unchanged. */
declare global {
  var __folioCanonicalFuzzErrors: { type: string; message: string }[] | undefined;
}

export const recordCanonicalFuzzError = (error: Error) => {
  globalThis.__folioCanonicalFuzzErrors?.push({ type: error.name, message: error.message });
};
