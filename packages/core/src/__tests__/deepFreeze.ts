/**
 * Deep-freeze for tests that prove a document is never written: every object
 * and array reachable from a value is frozen, and since maps and sets cannot
 * be frozen, their writing methods throw instead. A write in strict-mode code
 * then throws `TypeError` at the write.
 */

const refuseWrite = () => {
  throw new TypeError("A frozen document was written to.");
};

/**
 * Freeze everything reachable from `value`. Maps and sets cannot be frozen,
 * so their writing methods throw instead. Returns `value`.
 */
export const deepFreeze = <Value>(value: Value): Value => {
  freezeReachable(value, new Set());
  return value;
};

const freezeReachable = (value: unknown, seen: Set<unknown>): void => {
  if (
    typeof value !== "object" ||
    value === null ||
    seen.has(value) ||
    // Versions share what an edit left: a part frozen with an earlier one.
    Object.isFrozen(value)
  )
    return;
  seen.add(value);
  if (ArrayBuffer.isView(value) || value instanceof ArrayBuffer) return;
  if (value instanceof Map) {
    for (const [key, entry] of value) {
      freezeReachable(key, seen);
      freezeReachable(entry, seen);
    }
    Object.assign(value, {
      set: refuseWrite,
      delete: refuseWrite,
      clear: refuseWrite,
    });
  } else if (value instanceof Set) {
    for (const entry of value) freezeReachable(entry, seen);
    Object.assign(value, {
      add: refuseWrite,
      delete: refuseWrite,
      clear: refuseWrite,
    });
  } else {
    for (const key of Reflect.ownKeys(value))
      freezeReachable((value as Record<PropertyKey, unknown>)[key], seen);
  }
  Object.freeze(value);
};
