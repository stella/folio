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
