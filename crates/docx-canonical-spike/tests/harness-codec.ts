/** Test transport only. Product operation wire shapes and schemas are unchanged. */
import { TaggedError } from "better-result";

export class HarnessCodecError extends TaggedError("HarnessCodecError")<{ message: string }> {}

export type Encoded =
  | null
  | boolean
  | number
  | string
  | { tag: "undefined" | "hole" }
  | { tag: "date"; iso: string }
  | { tag: "uint8Array" | "arrayBuffer"; bytes: number[] }
  | { tag: "array"; items: Encoded[] }
  | { tag: "object" | "map"; entries: [string, Encoded][] };

const TAG_FIELDS = {
  undefined: ["tag"],
  hole: ["tag"],
  date: ["tag", "iso"],
  uint8Array: ["tag", "bytes"],
  arrayBuffer: ["tag", "bytes"],
  array: ["tag", "items"],
  object: ["tag", "entries"],
  map: ["tag", "entries"],
} as const satisfies Record<
  Exclude<Encoded, null | boolean | number | string>["tag"],
  readonly string[]
>;
const isTag = (tag: unknown): tag is keyof typeof TAG_FIELDS =>
  typeof tag === "string" && Object.hasOwn(TAG_FIELDS, tag);

const checkTagFields = (value: object, tag: keyof typeof TAG_FIELDS): void => {
  if (!Object.hasOwn(value, "tag"))
    throw new HarnessCodecError({ message: "A harness object needs a tag." });
  const fields = TAG_FIELDS[tag];
  const owned = Object.getOwnPropertyNames(value);
  if (owned.length !== fields.length || owned.some((key) => !fields.some((field) => field === key)))
    throw new HarnessCodecError({ message: `Unexpected fields in the ${tag} tag.` });
};

export const encodeTagged = (value: unknown): Encoded => {
  if (value === undefined) return { tag: "undefined" };
  if (value === null || typeof value === "string" || typeof value === "boolean") return value;
  if (typeof value === "number") {
    if (!Number.isFinite(value) || Object.is(value, -0))
      throw new HarnessCodecError({
        message: "Non-finite numbers and negative zero are outside this model codec.",
      });
    return value;
  }
  if (value instanceof Date) {
    if (!Number.isFinite(value.getTime()))
      throw new HarnessCodecError({ message: "Invalid model Date." });
    return { tag: "date", iso: value.toISOString() };
  }
  if (value instanceof Uint8Array) return { tag: "uint8Array", bytes: [...value] };
  if (value instanceof ArrayBuffer)
    return { tag: "arrayBuffer", bytes: [...new Uint8Array(value)] };
  if (value instanceof Map) {
    const entries: [string, Encoded][] = [];
    for (const [key, entry] of value) {
      if (typeof key !== "string")
        throw new HarnessCodecError({ message: "Model Map keys must be strings." });
      entries.push([key, encodeTagged(entry)]);
    }
    return { tag: "map", entries };
  }
  if (Array.isArray(value))
    return {
      tag: "array",
      items: Array.from({ length: value.length }, (_, index) =>
        Object.hasOwn(value, index) ? encodeTagged(value[index]) : { tag: "hole" },
      ),
    };
  if (typeof value === "object" && Object.getPrototypeOf(value) === Object.prototype)
    return {
      tag: "object",
      entries: Object.getOwnPropertyNames(value).map((key) => {
        const descriptor = Object.getOwnPropertyDescriptor(value, key);
        if (!descriptor || !("value" in descriptor))
          throw new HarnessCodecError({
            message: "Model accessors are outside the harness codec.",
          });
        return [key, encodeTagged(descriptor.value)] satisfies [string, Encoded];
      }),
    };
  throw new HarnessCodecError({ message: "Unsupported model value in harness codec." });
};

export const decodeTagged = (value: unknown): unknown => {
  if (value === null || typeof value === "string" || typeof value === "boolean") return value;
  if (typeof value === "number") {
    if (!Number.isFinite(value) || Object.is(value, -0))
      throw new HarnessCodecError({ message: "Invalid harness number." });
    return value;
  }
  if (Array.isArray(value))
    throw new HarnessCodecError({ message: "Raw arrays are not tagged harness values." });
  if (
    typeof value !== "object" ||
    value === null ||
    !("tag" in value) ||
    !Object.hasOwn(value, "tag") ||
    typeof value.tag !== "string"
  )
    throw new HarnessCodecError({ message: "A harness object needs a tag." });
  if (!isTag(value.tag))
    throw new HarnessCodecError({ message: `Unknown harness tag ${value.tag}.` });
  checkTagFields(value, value.tag);
  switch (value.tag) {
    case "undefined":
      return undefined;
    case "hole":
      throw new HarnessCodecError({ message: "A hole tag belongs only to an array slot." });
    case "date": {
      if (!("iso" in value) || typeof value.iso !== "string")
        throw new HarnessCodecError({ message: "Date ISO must be a string." });
      const date = new Date(value.iso);
      if (!Number.isFinite(date.getTime()) || date.toISOString() !== value.iso)
        throw new HarnessCodecError({ message: "Invalid ISO Date." });
      return date;
    }
    case "uint8Array":
    case "arrayBuffer": {
      if (!("bytes" in value) || !Array.isArray(value.bytes))
        throw new HarnessCodecError({ message: "Binary bytes must be an array." });
      if (
        !Array.from(value.bytes).every(
          (byte) => typeof byte === "number" && Number.isInteger(byte) && byte >= 0 && byte <= 255,
        )
      )
        throw new HarnessCodecError({ message: "Binary bytes must be integers in 0..255." });
      const bytes = Uint8Array.from(value.bytes);
      return value.tag === "uint8Array" ? bytes : bytes.buffer;
    }
    case "array": {
      if (!("items" in value) || !Array.isArray(value.items))
        throw new HarnessCodecError({ message: "Array items must be an array." });
      const result: unknown[] = [];
      result.length = value.items.length;
      for (const [index, item] of value.items.entries()) {
        if (typeof item === "object" && item !== null && "tag" in item && item.tag === "hole") {
          checkTagFields(item, "hole");
          continue;
        }
        result[index] = decodeTagged(item);
      }
      return result;
    }
    case "object":
    case "map": {
      if (!("entries" in value) || !Array.isArray(value.entries))
        throw new HarnessCodecError({ message: "Object and Map entries must be arrays." });
      const entries: [string, unknown][] = [];
      const keys = new Set<string>();
      for (const entry of value.entries) {
        if (!Array.isArray(entry) || entry.length !== 2)
          throw new HarnessCodecError({ message: "An entry must be a key/value pair." });
        if (typeof entry[0] !== "string")
          throw new HarnessCodecError({ message: "Harness object and Map keys must be strings." });
        if (keys.has(entry[0]))
          throw new HarnessCodecError({ message: "Harness entries repeat an owned key." });
        keys.add(entry[0]);
        entries.push([entry[0], decodeTagged(entry[1])]);
      }
      return value.tag === "map" ? new Map(entries) : Object.fromEntries(entries);
    }
    default:
      throw new HarnessCodecError({ message: "A harness object needs a tag." });
  }
};

/** Independent comparison copy: remove only capture symbols; keep all string-owned facts. */
export const withoutCaptureSymbols = (value: unknown): unknown => {
  if (value instanceof Date) return new Date(value);
  if (value instanceof Uint8Array) return value.slice();
  if (value instanceof ArrayBuffer) return value.slice(0);
  if (value instanceof Map)
    return new Map([...value].map(([key, entry]) => [key, withoutCaptureSymbols(entry)]));
  if (Array.isArray(value)) return value.map(withoutCaptureSymbols);
  if (typeof value === "object" && value !== null)
    return Object.fromEntries(
      Object.getOwnPropertyNames(value).map((key) => {
        const descriptor = Object.getOwnPropertyDescriptor(value, key);
        if (!descriptor || !("value" in descriptor))
          throw new HarnessCodecError({
            message: "Model accessors are outside the harness codec.",
          });
        return [key, withoutCaptureSymbols(descriptor.value)];
      }),
    );
  return value;
};

export const transportScope = (value: unknown) => {
  const seen = new WeakSet<object>();
  let symbols = 0;
  let sharedReferences = 0;
  let nonEnumerableStringFields = 0;
  let accessorFields = 0;
  const visit = (entry: unknown): void => {
    if (typeof entry !== "object" || entry === null) return;
    if (seen.has(entry)) {
      sharedReferences += 1;
      return;
    }
    seen.add(entry);
    symbols += Object.getOwnPropertySymbols(entry).length;
    if (entry instanceof Map) {
      for (const child of entry.values()) visit(child);
      return;
    }
    if (entry instanceof Date || entry instanceof Uint8Array || entry instanceof ArrayBuffer)
      return;
    for (const [key, descriptor] of Object.entries(Object.getOwnPropertyDescriptors(entry))) {
      if (!Array.isArray(entry) && !descriptor.enumerable) nonEnumerableStringFields += 1;
      if (!("value" in descriptor)) {
        accessorFields += 1;
        continue;
      }
      if (Array.isArray(entry) && key === "length") continue;
      visit(descriptor.value);
    }
  };
  visit(value);
  return {
    excludedCaptureSymbols: symbols,
    sharedReferences,
    nonEnumerableStringFields,
    accessorFields,
  };
};
