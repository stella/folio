import type { Document } from "./document";

/** `Value` with every reachable field, element and entry read-only. */
export type DeepReadonly<Value> = Value extends (...args: never[]) => unknown
  ? Value
  : Value extends ArrayBuffer | Date
    ? Value
    : Value extends ReadonlyMap<infer Key, infer Entry>
      ? ReadonlyMap<DeepReadonly<Key>, DeepReadonly<Entry>>
      : Value extends ReadonlySet<infer Entry>
        ? ReadonlySet<DeepReadonly<Entry>>
        : Value extends readonly (infer Element)[]
          ? readonly DeepReadonly<Element>[]
          : Value extends object
            ? { readonly [Key in keyof Value]: DeepReadonly<Value[Key]> }
            : Value;

/**
 * The editor's document model: in a canonical session the session's own
 * document, which the session never writes again once it has handed it out,
 * so a host may hold it as long as it likes; it must never write to it.
 */
export type ReadonlyDocument = DeepReadonly<Document>;
