/**
 * Own-property nested path helpers for locale JSON mutation.
 *
 * Locale key paths are untrusted input (CLI / sync tooling). Descend only
 * through own properties and refuse prototype-chain segments so a crafted
 * path cannot mutate `Object.prototype`.
 */

export type NestedObject = {
  [key: string]: string | NestedObject;
};

const UNSAFE_KEYS = new Set(["__proto__", "constructor", "prototype"]);

const assertSafePath = (path: string, parts: string[], action: "set" | "delete" | "get"): void => {
  if (parts.some((part) => UNSAFE_KEYS.has(part))) {
    throw new Error(`Refusing to ${action} unsafe key path: ${path}`);
  }
};

/** Read a dotted path, ignoring inherited properties. */
export const getNestedValue = (
  obj: NestedObject,
  path: string,
): string | NestedObject | undefined => {
  const parts = path.split(".");
  assertSafePath(path, parts, "get");
  let current: string | NestedObject = obj;

  for (const part of parts) {
    if (typeof current === "string" || !Object.hasOwn(current, part)) {
      return undefined;
    }
    const next = current[part];
    if (next === undefined) {
      return undefined;
    }
    current = next;
  }

  return current;
};

/** Set a dotted path, creating only own enumerable data properties. */
export const setNestedValue = (
  obj: NestedObject,
  path: string,
  value: string | NestedObject,
): void => {
  const parts = path.split(".");
  assertSafePath(path, parts, "set");
  const leaf = parts.at(-1);
  if (!leaf) {
    return;
  }

  let current: NestedObject = obj;
  for (const part of parts.slice(0, -1)) {
    if (!Object.hasOwn(current, part) || typeof current[part] !== "object" || current[part] === null) {
      const child: NestedObject = {};
      Object.defineProperty(current, part, {
        value: child,
        writable: true,
        enumerable: true,
        configurable: true,
      });
      current = child;
      continue;
    }
    // SAFETY: Object.hasOwn + typeof object + non-null narrows to NestedObject.
    current = current[part] as NestedObject;
  }

  Object.defineProperty(current, leaf, {
    value,
    writable: true,
    enumerable: true,
    configurable: true,
  });
};

/** Delete a dotted path and prune empty own-property parents. */
export const deleteNestedValue = (obj: NestedObject, path: string): void => {
  const parts = path.split(".");
  assertSafePath(path, parts, "delete");
  const leaf = parts.at(-1);
  if (!leaf) {
    return;
  }

  const stack: NestedObject[] = [obj];
  for (let i = 0; i < parts.length - 1; i++) {
    const key = parts[i]!; // SAFETY: i < parts.length - 1
    const parent = stack.at(-1);
    if (!parent || !Object.hasOwn(parent, key)) {
      return;
    }
    const next = parent[key];
    if (next === undefined || typeof next === "string") {
      return;
    }
    stack.push(next);
  }

  const target = stack.at(-1);
  if (!target || !Object.hasOwn(target, leaf)) {
    return;
  }
  Reflect.deleteProperty(target, leaf);

  for (let i = stack.length - 1; i > 0; i--) {
    const child = stack.at(i);
    const parent = stack.at(i - 1);
    const key = parts.at(i - 1);
    if (!child || !parent || key === undefined) {
      break;
    }
    if (Object.keys(child).length === 0) {
      Reflect.deleteProperty(parent, key);
      continue;
    }
    break;
  }
};
