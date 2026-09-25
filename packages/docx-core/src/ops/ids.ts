/**
 * Paragraph ids across a whole package.
 *
 * `w14:paraId` is unique across every part of a package, not per story: body,
 * headers, footers, notes, comments and the text boxes inside any of them
 * share one space. The walk is structural, so a story the model gains is
 * covered without a list to keep in step.
 */

const isPlainRecord = (value: object): boolean => {
  const prototype: unknown = Object.getPrototypeOf(value);
  return prototype === Object.prototype || prototype === null;
};

const collect = (value: unknown, out: string[]): void => {
  if (typeof value !== "object" || value === null) {
    return;
  }
  if (Array.isArray(value)) {
    for (const item of value) collect(item, out);
    return;
  }
  if (value instanceof Map) {
    for (const item of value.values()) collect(item, out);
    return;
  }
  // Bytes, dates and other built-ins hold no paragraphs.
  if (!isPlainRecord(value)) {
    return;
  }
  const entries = Object.entries(value);
  const isParagraph = entries.some(([key, field]) => key === "type" && field === "paragraph");
  for (const [key, field] of entries) {
    if (isParagraph && key === "paraId" && typeof field === "string") {
      out.push(field);
    }
    collect(field, out);
  }
};

/** Every `paraId` of every paragraph in a value, nested ones included, in document order. */
export const paragraphIdsIn = (value: unknown): string[] => {
  const out: string[] = [];
  collect(value, out);
  return out;
};

/**
 * The key two spellings of one id share. `ST_LongHexNumber` is hex, so
 * `0000abcd` and `0000ABCD` are the same id.
 */
export const idKey = (id: string): string => id.toUpperCase();

/** How many paragraphs carry each id, by {@link idKey}. */
export const countIds = (ids: readonly string[]): Map<string, number> => {
  const counts = new Map<string, number>();
  for (const id of ids) counts.set(idKey(id), (counts.get(idKey(id)) ?? 0) + 1);
  return counts;
};

/** Whether `incoming` repeats an id of `existing` or of itself, by {@link idKey}. */
export const collides = (
  existing: ReadonlyMap<string, number>,
  incoming: readonly string[],
): boolean => {
  const seen = new Set<string>();
  return incoming.some((id) => {
    const key = idKey(id);
    const repeated = seen.has(key) || (existing.get(key) ?? 0) > 0;
    seen.add(key);
    return repeated;
  });
};

/** `ST_LongHexNumber` as `w14:paraId` uses it: eight hex digits below `0x80000000`. */
const PARA_ID_PATTERN = /^[0-7][0-9A-Fa-f]{7}$/u;
/** Zero is reserved: it means "no id". */
const RESERVED_PARA_ID_PATTERN = /^0{8}$/u;

/** Whether an id can name a paragraph an operation creates. */
export const isParaId = (id: string): boolean =>
  PARA_ID_PATTERN.test(id) && !RESERVED_PARA_ID_PATTERN.test(id);
