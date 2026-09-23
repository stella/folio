import { panic } from "better-result";
import type { Attrs, Mark, MarkType, Schema } from "prosemirror-model";

import type { MarkFactory } from "../extensions/marks/markUtils";

/**
 * Serialize one attr value so that equal keys mean interchangeable values.
 *
 * Numbers keep `-0`, `NaN` and the infinities distinct, and `undefined` stays
 * distinct from `null` and from an absent key. Returns `null` for anything that
 * is not plain data (class instances, functions, symbols, sparse arrays); the
 * caller then builds an unshared mark.
 */
const valueKey = (value: unknown): string | null => {
  switch (typeof value) {
    case "string":
      return JSON.stringify(value);
    case "number":
      return Object.is(value, -0) ? "n-0" : `n${value}`;
    case "boolean":
      return value ? "t" : "f";
    case "undefined":
      return "u";
    case "object":
      return value === null ? "z" : objectKey(value);
    default:
      return null;
  }
};

const objectKey = (value: object): string | null => {
  if (Array.isArray(value)) {
    const parts: string[] = [];
    for (let index = 0; index < value.length; index += 1) {
      if (!(index in value)) {
        return null;
      }
      const part = valueKey(value[index]);
      if (part === null) {
        return null;
      }
      parts.push(part);
    }
    return `[${parts.join(",")}]`;
  }
  const prototype: unknown = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) {
    return null;
  }
  // Key order is part of the key: an interned mark hands its nested objects to
  // every run that shares it, and serialized output follows their key order.
  const parts: string[] = [];
  for (const [key, entry] of Object.entries(value)) {
    const part = valueKey(entry);
    if (part === null) {
      return null;
    }
    parts.push(`${JSON.stringify(key)}:${part}`);
  }
  return `{${parts.join(",")}}`;
};

/**
 * Key a mark by its type and the attrs `MarkType.create` would compute.
 *
 * Only declared attrs are read, and a top-level `undefined` is keyed like an
 * absent attr, because `create` ignores undeclared keys and fills both from
 * the attr's default.
 */
const markKey = (
  type: MarkType,
  attrNames: readonly string[],
  attrs: Attrs | null,
): string | null => {
  const parts: string[] = [type.name];
  for (const name of attrNames) {
    const value: unknown = attrs?.[name];
    const part = value === undefined ? "d" : valueKey(value);
    if (part === null) {
      return null;
    }
    parts.push(part);
  }
  return parts.join("\u0000");
};

/**
 * Build marks that share one instance per distinct (type, attrs) value.
 *
 * A document repeats a small set of run formattings across thousands of runs,
 * so one conversion builds each distinct mark once. Marks are immutable, and
 * ProseMirror compares them by value, so sharing an instance changes nothing a
 * caller can observe except the saved allocation. Create one interner per
 * conversion; it holds every mark it built until it is dropped.
 */
export const createMarkInterner = (schema: Schema): MarkFactory => {
  const marks = new Map<string, Mark>();
  const attrNamesByType = new Map<MarkType, readonly string[]>();
  return (typeName, attrs) => {
    const type = schema.marks[typeName];
    if (type === undefined) {
      panic(`Unknown mark type ${typeName}.`);
    }
    let attrNames = attrNamesByType.get(type);
    if (attrNames === undefined) {
      attrNames = Object.keys(type.spec.attrs ?? {});
      attrNamesByType.set(type, attrNames);
    }
    const key = markKey(type, attrNames, attrs ?? null);
    if (key === null) {
      return type.create(attrs);
    }
    const cached = marks.get(key);
    if (cached !== undefined) {
      return cached;
    }
    const mark = type.create(attrs);
    marks.set(key, mark);
    return mark;
  };
};
