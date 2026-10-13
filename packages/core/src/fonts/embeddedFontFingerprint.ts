/** Immutable editor-owned lookups share content fingerprints without retaining disposed editors. */
const fingerprints = new WeakMap<ReadonlyMap<string, string>, string>();

/** Callers replace owned maps on font changes; mutating a published map is unsupported. */
export const embeddedFontFingerprint = (families: ReadonlyMap<string, string>): string => {
  const cached = fingerprints.get(families);
  if (cached !== undefined) return cached;

  const entries = [...families].sort(([left], [right]) => {
    if (left < right) return -1;
    if (left > right) return 1;
    return 0;
  });
  const fingerprint = JSON.stringify(entries);
  fingerprints.set(families, fingerprint);
  return fingerprint;
};
