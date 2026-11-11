import { toRaw } from "vue";

import type { Document } from "@stll/folio-core/types/document";

/**
 * A host's document with every reactive proxy reachable from it replaced by
 * its raw object, so the editor's model can be copied: `structuredClone`
 * refuses a proxy. `toRaw` alone unwraps only the object it is given, and a
 * host can hold a proxy deeper down, as a spread of a reactive document does.
 *
 * Nothing is written: a document that reaches no proxy is returned as it is,
 * and an object on the way to a proxy is rebuilt around the raw object, with
 * its own property descriptors (symbol-keyed ones included) carried over.
 */
export const rawDocument = <Value extends Document>(document: Value): Value =>
  // SAFETY: `rawData` keeps every value's shape; it only swaps proxies for
  // the raw objects they wrap.
  rawData(document, new Map()) as Value;

const rawData = (value: unknown, seen: Map<object, unknown>): unknown => {
  if (typeof value !== "object" || value === null) return value;
  const raw: object = toRaw(value);
  const known = seen.get(raw);
  if (known !== undefined) return known;
  // A cycle back to an object being walked keeps the object itself.
  seen.set(raw, raw);
  const result = rebuild(raw, seen);
  seen.set(raw, result);
  return result;
};

const rebuild = (raw: object, seen: Map<object, unknown>): object => {
  if (ArrayBuffer.isView(raw) || raw instanceof ArrayBuffer) return raw;
  if (raw instanceof Map) {
    let changed = false;
    const entries: [unknown, unknown][] = [];
    for (const [key, entry] of raw) {
      const nextKey = rawData(key, seen);
      const nextEntry = rawData(entry, seen);
      if (nextKey !== key || nextEntry !== entry) changed = true;
      entries.push([nextKey, nextEntry]);
    }
    return changed ? new Map(entries) : raw;
  }
  if (raw instanceof Set) {
    let changed = false;
    const entries: unknown[] = [];
    for (const entry of raw) {
      const next = rawData(entry, seen);
      if (next !== entry) changed = true;
      entries.push(next);
    }
    return changed ? new Set(entries) : raw;
  }
  if (Array.isArray(raw)) {
    let changed = false;
    const items = raw.map((item: unknown) => {
      const next = rawData(item, seen);
      if (next !== item) changed = true;
      return next;
    });
    return changed ? items : raw;
  }
  const descriptors = Object.getOwnPropertyDescriptors(raw);
  let changed = false;
  for (const key of Reflect.ownKeys(descriptors)) {
    const descriptor = Reflect.get(descriptors, key);
    if (!isDataDescriptor(descriptor)) continue;
    const next = rawData(descriptor.value, seen);
    if (next === descriptor.value) continue;
    changed = true;
    descriptor.value = next;
  }
  return changed
    ? Object.create(Reflect.getPrototypeOf(raw), descriptors)
    : raw;
};

const isDataDescriptor = (
  descriptor: unknown,
): descriptor is PropertyDescriptor & { value: unknown } =>
  typeof descriptor === "object" &&
  descriptor !== null &&
  "value" in descriptor;
