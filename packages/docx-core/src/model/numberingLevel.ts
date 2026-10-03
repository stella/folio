/** A w:ilvl outside the nine numbering levels names no level. */
const NUMBERING_LEVELS = new Set([0, 1, 2, 3, 4, 5, 6, 7, 8]);

/** Resolve only defined numbering levels; preserve unsupported source values elsewhere. */
export const isNumberingLevel = (value: number): boolean => NUMBERING_LEVELS.has(value);
