type ModelSnapshot = {
  value: Record<string, unknown>;
  keys: string[];
  children: unknown[];
  descendants: (ModelSnapshot | undefined)[];
  arrayLength: number | undefined;
  version: object;
  readToken: object;
};

/** Exact version keys for mutable model records and arrays, without serialization. */
export const createModelVersionTracker = () => {
  const snapshots = new WeakMap<object, ModelSnapshot>();
  let readToken = {};

  const isRecord = (value: unknown): value is Record<string, unknown> =>
    value !== null && typeof value === "object";

  const read = (value: unknown, descendant?: ModelSnapshot): unknown => {
    if (!isRecord(value)) return value;
    // Retain edges as well as versions: a warm tree needs no WeakMap lookup at
    // every child. A shared record is inspected once per read, never once per
    // alias. A later read still checks every field for in-place mutations.
    const previous = descendant?.value === value ? descendant : snapshots.get(value);
    if (previous?.readToken === readToken) return previous.version;
    const arrayLength = Array.isArray(value) ? value.length : undefined;
    let changed = previous === undefined || previous.arrayLength !== arrayLength;
    let keys = changed ? [] : (previous?.keys ?? []);
    let children = changed ? [] : (previous?.children ?? []);
    let descendants = changed ? [] : (previous?.descendants ?? []);
    let index = 0;
    for (const key of Object.keys(value)) {
      const field = value[key];
      const childRecord = field !== null && typeof field === "object";
      const child = childRecord ? read(field, previous?.descendants[index]) : field;
      if (
        !changed &&
        previous &&
        (index >= previous.keys.length ||
          previous.keys[index] !== key ||
          !Object.is(previous.children[index], child))
      ) {
        keys = previous.keys.slice(0, index);
        children = previous.children.slice(0, index);
        descendants = previous.descendants.slice(0, index);
        changed = true;
      }
      if (changed) {
        keys.push(key);
        children.push(child);
        descendants.push(childRecord ? snapshots.get(field) : undefined);
      }
      index += 1;
    }
    if (!changed && previous && index !== previous.keys.length) {
      keys = previous.keys.slice(0, index);
      children = previous.children.slice(0, index);
      descendants = previous.descendants.slice(0, index);
      changed = true;
    }
    if (!changed && previous) {
      previous.readToken = readToken;
      return previous.version;
    }
    const version = {};
    if (previous) {
      Object.assign(previous, { keys, children, descendants, arrayLength, version, readToken });
    } else {
      snapshots.set(value, { value, keys, children, descendants, arrayLength, version, readToken });
    }
    return version;
  };

  return (value: unknown) => {
    readToken = {};
    return read(value);
  };
};
