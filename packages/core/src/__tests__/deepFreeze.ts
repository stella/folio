/**
 * Deep-freeze for tests that prove a document is never written: every object
 * and array reachable from a value is frozen, and since maps and sets cannot
 * be frozen, their writing methods throw instead. A write in strict-mode code
 * then throws `TypeError` at the write, and the value can still be copied.
 */

const refuseWrite = (): never => {
  throw new TypeError("A frozen document was written to.");
};

// A frozen map or set keeps its internal data, so `structuredClone` still
// copies it; only its prototype changes, to one whose writers throw.
class FrozenMap<Key, Entry> extends Map<Key, Entry> {
  override set(): never {
    return refuseWrite();
  }
  override delete(): never {
    return refuseWrite();
  }
  override clear(): never {
    return refuseWrite();
  }
}

class FrozenSet<Entry> extends Set<Entry> {
  override add(): never {
    return refuseWrite();
  }
  override delete(): never {
    return refuseWrite();
  }
  override clear(): never {
    return refuseWrite();
  }
}

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
    Object.setPrototypeOf(value, FrozenMap.prototype);
  } else if (value instanceof Set) {
    for (const entry of value) freezeReachable(entry, seen);
    Object.setPrototypeOf(value, FrozenSet.prototype);
  } else {
    for (const key of Reflect.ownKeys(value))
      freezeReachable((value as Record<PropertyKey, unknown>)[key], seen);
  }
  Object.freeze(value);
};
