/**
 * Structural equality over model records: plain objects, arrays and
 * primitives. A key holding `undefined` is the same as an absent key, which is
 * how the value reads after a JSON round-trip.
 */

/** The keys of a record a comparison leaves out. */
type SkippedKeys = (entries: readonly [string, unknown][]) => ReadonlySet<string>;

const NOTHING: ReadonlySet<string> = new Set();

const equal = (left: unknown, right: unknown, skipped: SkippedKeys): boolean => {
  if (Object.is(left, right)) {
    return true;
  }
  if (typeof left !== "object" || typeof right !== "object" || left === null || right === null) {
    return false;
  }
  if (Array.isArray(left) || Array.isArray(right)) {
    if (!Array.isArray(left) || !Array.isArray(right) || left.length !== right.length) {
      return false;
    }
    return left.every((item, index) => equal(item, right[index], skipped));
  }
  const fields = (record: object): [string, unknown][] => {
    const entries = Object.entries(record);
    const skip = skipped(entries);
    return entries.filter(([key, value]) => value !== undefined && !skip.has(key));
  };
  const leftEntries = fields(left);
  const rightEntries = new Map(fields(right));
  if (leftEntries.length !== rightEntries.size) {
    return false;
  }
  return leftEntries.every(
    ([key, value]) => rightEntries.has(key) && equal(value, rightEntries.get(key), skipped),
  );
};

export const structurallyEqual = (left: unknown, right: unknown): boolean =>
  equal(left, right, () => NOTHING);

/**
 * Paragraph fields a layout or numbering pass recomputes. They describe the
 * content rather than being it, so they never make an operation's
 * precondition stale.
 */
const DERIVED_PARAGRAPH_FIELDS: ReadonlySet<string> = new Set([
  "listRendering",
  "renderedPageBreakBefore",
]);

const derivedFields: SkippedKeys = (entries) =>
  entries.some(([key, value]) => key === "type" && value === "paragraph")
    ? DERIVED_PARAGRAPH_FIELDS
    : NOTHING;

/**
 * The equality an operation's precondition is checked under: structural,
 * with the fields a relayout or renumbering recomputes left out. An undo is
 * refused as stale when the content it expects has changed, never because
 * numbering or pagination was recomputed since.
 */
export const equalForStaleness = (left: unknown, right: unknown): boolean =>
  equal(left, right, derivedFields);
