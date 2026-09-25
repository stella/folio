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

/** How many paragraphs carry each id. */
export const countIds = (ids: readonly string[]): Map<string, number> => {
  const counts = new Map<string, number>();
  for (const id of ids) counts.set(id, (counts.get(id) ?? 0) + 1);
  return counts;
};
