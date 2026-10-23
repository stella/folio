import { boundedContentIdentity } from "../utils/boundedContentIdentity";

/** Immutable editor-owned lookups reuse cache-local identities without retaining disposed editors. */
const fingerprints = new WeakMap<ReadonlyMap<string, string>, string>();

/** Callers replace owned maps on font changes; mutating a published map is unsupported. */
export const embeddedFontFingerprint = (
  families: ReadonlyMap<string, string>,
): string => {
  const cached = fingerprints.get(families);
  if (cached !== undefined) return cached;

  const entries = [...families].sort(([left], [right]) => {
    if (left < right) return -1;
    if (left > right) return 1;
    return 0;
  });
  const fingerprint = boundedContentIdentity(JSON.stringify(entries));
  fingerprints.set(families, fingerprint);
  return fingerprint;
};
