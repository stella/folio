/**
 * Block ids name a block, they do not describe it: a conversion that keeps
 * blocks from an earlier one hands out other ids than a conversion from
 * scratch. Equivalence tests compare the two after renaming every block id by
 * the order a walk over the value meets it, so they still catch two blocks
 * sharing an id, a reference to the wrong block, and anything else that is
 * not just a different name.
 */

/** A block id handed out by a conversion, with what a derived id appends. */
const BLOCK_ID = /^(block-\d+)(.*)$/s;
/** Fields that hold a block id: blocks, table rows and cells, and fragments. */
const ID_KEYS = new Set(["id", "blockId"]);

const isPlainObject = (value: object): value is Record<string, unknown> => {
  const prototype = Object.getPrototypeOf(value);
  return prototype === Object.prototype || prototype === null;
};

type Walk = {
  readonly onId: (id: string, key: string) => string;
  /** Copies of the objects walked: met again, an object is not walked again. */
  readonly seen: WeakMap<object, unknown>;
  /** Walk an object each time it is met, not only the first time. */
  readonly everyTime: boolean;
};

/**
 * Walk `value` in a fixed order: arrays and maps in order, object fields by
 * name, so two values equal up to block ids are walked alike.
 */
function walk(value: unknown, at: Walk): unknown {
  if (typeof value !== "object" || value === null) return value;
  const known = at.seen.get(value);
  if (known !== undefined) return known;
  const copy = copyOf(value, at);
  if (at.everyTime) at.seen.delete(value);
  return copy;
}

function copyOf(value: object, at: Walk): unknown {
  const { seen } = at;
  if (Array.isArray(value)) {
    const copy: unknown[] = [];
    seen.set(value, copy);
    for (const item of value) copy.push(walk(item, at));
    return copy;
  }
  if (value instanceof Map) {
    const copy = new Map<unknown, unknown>();
    seen.set(value, copy);
    for (const [key, item] of value) copy.set(walk(key, at), walk(item, at));
    return copy;
  }
  if (value instanceof Set) {
    const copy = new Set<unknown>();
    seen.set(value, copy);
    for (const item of value) copy.add(walk(item, at));
    return copy;
  }
  // Anything else (a document node, a class instance) is not a layout value.
  if (!isPlainObject(value)) return value;
  const copy: Record<string, unknown> = {};
  seen.set(value, copy);
  for (const key of Object.keys(value).sort()) {
    const field = value[key];
    copy[key] =
      ID_KEYS.has(key) && typeof field === "string" && BLOCK_ID.test(field)
        ? at.onId(field, key)
        : walk(field, at);
  }
  return copy;
}

/**
 * A copy of `value` (blocks, a layout, or anything holding them) with each
 * block id renamed `block-1`, `block-2`, ... in the order the walk meets it.
 * An id derived from another (`block-4:frame`) keeps its suffix on the
 * other's new name.
 */
export function renameBlockIds<T>(value: T): T {
  const names = new Map<string, string>();
  return walk(value, {
    onId: (id) => {
      const [, base = id, suffix = ""] = BLOCK_ID.exec(id) ?? [];
      let name = names.get(base);
      if (name === undefined) {
        name = `block-${names.size + 1}`;
        names.set(base, name);
      }
      return `${name}${suffix}`;
    },
    seen: new WeakMap(),
    everyTime: false,
  }) as T;
}

/**
 * The id of every block, row and cell in `value`, nested ones included, in
 * walk order. A block met twice counts twice: it is two blocks with one id.
 */
export function blockIdsIn(value: unknown): string[] {
  const ids: string[] = [];
  walk(value, {
    onId: (id, key) => {
      if (key === "id") ids.push(id);
      return id;
    },
    seen: new WeakMap(),
    everyTime: true,
  });
  return ids;
}
