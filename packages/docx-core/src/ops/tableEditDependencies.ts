/** Content removed with a cell may own records outside that cell. */
const DEPENDENT_TYPES = new Set([
  "commentRangeStart",
  "commentRangeEnd",
  "commentReference",
  "bookmarkStart",
  "bookmarkEnd",
  "insertion",
  "deletion",
  "moveFrom",
  "moveTo",
  "preservedXml",
  "contentControl",
]);

/** Walk nested tables and inline containers as well as direct paragraphs. */
export const hasTableCellDependencies = (value: unknown): boolean => {
  if (typeof value !== "object" || value === null) return false;
  if (Array.isArray(value)) return value.some(hasTableCellDependencies);
  const type: unknown = Reflect.get(value, "type");
  if (typeof type === "string" && DEPENDENT_TYPES.has(type)) return true;
  for (const [key, field] of Object.entries(value)) {
    if (
      [
        "propertyChanges",
        "structuralChange",
        "pPrMark",
        "bookmarks",
        "preserved",
        "carrierStack",
      ].includes(key) &&
      field !== undefined
    )
      return true;
    if (hasTableCellDependencies(field)) return true;
  }
  return false;
};
