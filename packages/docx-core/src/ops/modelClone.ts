import { panic } from "better-result";

/** Clone model records while retaining opaque, enumerable in-memory bindings. */
type CopyPrivateBindingsOptions = { source: unknown; target: unknown; seen: WeakSet<object> };
const copyPrivateBindings = ({ source, target, seen }: CopyPrivateBindingsOptions): void => {
  if (
    typeof source !== "object" ||
    source === null ||
    typeof target !== "object" ||
    target === null ||
    seen.has(source)
  )
    return;
  seen.add(source);
  for (const key of Object.getOwnPropertySymbols(source)) {
    const descriptor = Object.getOwnPropertyDescriptor(source, key);
    if (descriptor?.enumerable && "value" in descriptor)
      Object.defineProperty(target, key, descriptor);
  }
  if (source instanceof Map && target instanceof Map) {
    const targets = target.entries();
    for (const [key, value] of source) {
      const entry = targets.next().value;
      if (!entry) panic("A model clone must preserve map entries.");
      copyPrivateBindings({ source: key, target: entry[0], seen });
      copyPrivateBindings({ source: value, target: entry[1], seen });
    }
    return;
  }
  if (source instanceof Set && target instanceof Set) {
    const targets = target.values();
    for (const value of source) {
      const entry = targets.next();
      if (entry.done) panic("A model clone must preserve set entries.");
      copyPrivateBindings({ source: value, target: entry.value, seen });
    }
    return;
  }
  if (
    ArrayBuffer.isView(source) ||
    source instanceof ArrayBuffer ||
    source instanceof Date ||
    source instanceof RegExp
  )
    return;
  for (const key of Object.keys(source))
    copyPrivateBindings({
      source: Reflect.get(source, key),
      target: Reflect.get(target, key),
      seen,
    });
};

/** Symbols carry capabilities owned by their producer; their handles retain identity.
 * String-key model data, binary buffers, aliases and cycles are deeply cloned.
 * JSON payloads have no bindings to transfer.
 */
export const cloneModel = <T>(value: T) => {
  const cloned = structuredClone(value);
  copyPrivateBindings({ source: value, target: cloned, seen: new WeakSet() });
  return cloned;
};
