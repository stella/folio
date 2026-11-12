/**
 * Every object of document data reachable from `value`. Symbol-keyed handles
 * are left out: they carry a paragraph's parse provenance, which the copy
 * owner hands on to the copy on purpose, and hold no document content.
 */
const reachable = (value: unknown, seen = new Set<object>()): Set<object> => {
  if (typeof value !== "object" || value === null || seen.has(value))
    return seen;
  seen.add(value);
  if (value instanceof Map) {
    for (const [key, entry] of value) {
      reachable(key, seen);
      reachable(entry, seen);
    }
  } else if (value instanceof Set) {
    for (const entry of value) reachable(entry, seen);
  } else if (!(ArrayBuffer.isView(value) || value instanceof ArrayBuffer)) {
    for (const entry of Object.values(value)) reachable(entry, seen);
  }
  return seen;
};

/** The objects `result` shares with `input`. */
export const sharedObjects = (input: unknown, result: unknown): object[] => {
  const inputs = reachable(input);
  return [...reachable(result)].filter((object) => inputs.has(object));
};
