import type { DocumentFontFacts } from "../src/fonts/documentFontFacts";

type StructuralFontFacts = {
  readonly cacheIdentity: string;
  fontInfoOf(name: string): undefined;
};

export const fontFactsRequireOwnership =
  true satisfies StructuralFontFacts extends DocumentFontFacts ? false : true;
