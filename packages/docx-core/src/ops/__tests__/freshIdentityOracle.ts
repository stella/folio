import { deepStrictEqual } from "node:assert/strict";

import type { Document } from "../../model/document";
import { IDENTITY_SPACES, packageIdentityKeys, slotKey } from "../ids";
import type { NewIds } from "../types";

type FreshIdentityOracleOptions = {
  actual: Document;
  expected: Document;
  original: Document;
  allocated: NewIds;
};

/** One document-wide bijection; original identities and every other field stay exact. */
export const assertFreshIdentityEquivalent = ({
  actual,
  expected,
  original,
  allocated,
}: FreshIdentityOracleOptions): void => {
  const source = new Set(packageIdentityKeys(original.package));
  const allocation = new Map<string, number>();
  for (const space of Object.values(IDENTITY_SPACES)) {
    for (const [index, id] of (allocated[space] ?? []).entries()) {
      const key = slotKey({ space, id });
      deepStrictEqual(source.has(key), false, "Fresh allocations cannot reuse source identities.");
      deepStrictEqual(allocation.has(key), false, "Fresh allocations must be unique.");
      allocation.set(key, index);
    }
  }
  const canonical = (document: Document): unknown => {
    const fresh = new Map<string, string>();
    const previous = new Map<string, number>();
    for (const key of packageIdentityKeys(document.package)) {
      if (source.has(key)) continue;
      const rank = allocation.get(key);
      deepStrictEqual(
        rank === undefined,
        false,
        "Every fresh identity must come from the allocation.",
      );
      if (rank === undefined) throw new Error("A checked allocation must exist.");
      deepStrictEqual(
        fresh.has(key),
        false,
        "Every fresh identity must be unique in the document.",
      );
      const space = key.split(":").at(0) ?? "";
      deepStrictEqual(
        rank > (previous.get(space) ?? -1),
        true,
        "Fresh identities follow allocation order in the document.",
      );
      previous.set(space, rank);
      fresh.set(key, `fresh:${space}:${fresh.size}`);
    }
    const walk = (value: unknown, heldBy?: string): unknown => {
      if (Array.isArray(value)) return value.map((item) => walk(item, heldBy));
      if (value instanceof Map) return new Map([...value].map(([key, item]) => [key, walk(item)]));
      if (value === null || typeof value !== "object") return value;
      if (Object.getPrototypeOf(value) !== Object.prototype) return value;
      const identitySpace = () => {
        if (heldBy === "info" && typeof Reflect.get(value, "author") === "string")
          return IDENTITY_SPACES.REVISION;
        if (typeof Reflect.get(value, "sdtType") === "string") return IDENTITY_SPACES.CONTROL;
        return undefined;
      };
      const space = identitySpace();
      return Object.fromEntries(
        Object.entries(value).map(([key, field]) => [
          key,
          key === "id" && typeof field === "number" && space !== undefined
            ? (fresh.get(slotKey({ space, id: field })) ?? field)
            : walk(field, key),
        ]),
      );
    };
    return walk(document);
  };
  deepStrictEqual(canonical(actual), canonical(expected));
};
