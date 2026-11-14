/**
 * Records a removed column cell cannot take with it: comments and moves pair
 * with records outside the cell, and captured markup is opaque. Bookmark
 * boundaries, content controls and the review records of the cell's own
 * words go with the cell, as they go with a removed row.
 */
const REACHING_TYPES = new Set([
  "commentRangeStart",
  "commentRangeEnd",
  "commentReference",
  "moveFrom",
  "moveTo",
  "moveFromRangeStart",
  "moveFromRangeEnd",
  "moveToRangeStart",
  "moveToRangeEnd",
  "preservedXml",
  "preservedBlock",
  "preservedInline",
]);

/** Walk nested tables and inline containers as well as direct paragraphs. */
export const reachesPastRemovedCell = (value: unknown): boolean => {
  if (typeof value !== "object" || value === null) return false;
  if (Array.isArray(value)) return value.some(reachesPastRemovedCell);
  const type: unknown = Reflect.get(value, "type");
  if (typeof type === "string" && REACHING_TYPES.has(type)) return true;
  for (const [key, field] of Object.entries(value)) {
    if (key === "preserved" && field !== undefined) return true;
    if (reachesPastRemovedCell(field)) return true;
  }
  return false;
};

/**
 * Whether a bookmark starts or ends in the removed cells while its other
 * boundary stays: removing one boundary would leave the other unpaired.
 */
export const strandsABookmark = (cells: readonly unknown[]): boolean => {
  const starts = new Set<unknown>();
  const ends = new Set<unknown>();
  const walk = (value: unknown): void => {
    if (typeof value !== "object" || value === null) return;
    if (Array.isArray(value)) {
      for (const item of value) walk(item);
      return;
    }
    const type: unknown = Reflect.get(value, "type");
    if (type === "bookmarkStart") starts.add(Reflect.get(value, "id"));
    if (type === "bookmarkEnd") ends.add(Reflect.get(value, "id"));
    for (const field of Object.values(value)) walk(field);
  };
  walk(cells);
  return (
    [...starts].some((id) => !ends.has(id)) ||
    [...ends].some((id) => !starts.has(id))
  );
};
