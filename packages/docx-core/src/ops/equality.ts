/**
 * Structural equality over model records: plain objects, arrays and
 * primitives. A key holding `undefined` is the same as an absent key, which is
 * how the value reads after a JSON round-trip.
 */
export const structurallyEqual = (left: unknown, right: unknown): boolean => {
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
    return left.every((item, index) => structurallyEqual(item, right[index]));
  }
  const leftEntries = Object.entries(left).filter(([, value]) => value !== undefined);
  const rightEntries = new Map(Object.entries(right).filter(([, value]) => value !== undefined));
  if (leftEntries.length !== rightEntries.size) {
    return false;
  }
  return leftEntries.every(
    ([key, value]) => rightEntries.has(key) && structurallyEqual(value, rightEntries.get(key)),
  );
};
