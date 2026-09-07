/**
 * JSON with object keys in a fixed order, for comparing two values that were
 * built by different code paths.
 *
 * `JSON.stringify` writes keys in insertion order, so two objects holding the
 * same properties compare unequal whenever one was assembled in a different
 * sequence — which is what happens whenever a parsed record is compared with a
 * rebuilt one. `undefined` members are omitted, so an absent property and one
 * explicitly set to `undefined` read alike.
 */
export const canonicalJson = (value: unknown): string => {
  if (value === null || value === undefined) {
    return "null";
  }
  if (Array.isArray(value)) {
    return `[${value.map(canonicalJson).join(",")}]`;
  }
  if (typeof value !== "object") {
    return JSON.stringify(value) ?? "null";
  }
  const entries = Object.entries(value)
    .filter(([, entry]) => entry !== undefined)
    .toSorted(([left], [right]) => left.localeCompare(right));
  return `{${entries
    .map(([key, entry]) => `${JSON.stringify(key)}:${canonicalJson(entry)}`)
    .join(",")}}`;
};
