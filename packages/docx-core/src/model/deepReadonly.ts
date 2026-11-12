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

declare const writableRecord: unique symbol;

/**
 * Every record type of the model carries this field, so its `DeepReadonly`
 * form is not assignable to it: TypeScript ignores `readonly` on object
 * fields when it compares types and checks it only on arrays, so a read-only
 * record would otherwise pass to any function that takes the writable record
 * and writes it. The field's type is a mutable array, which `DeepReadonly`
 * turns into a read-only one, and a read-only array is the one read-only
 * value TypeScript refuses where a writable one is expected. The field is
 * optional and never set: a writable record still passes where its read-only
 * form is expected, and an object literal needs nothing extra. A writer copies
 * a read-only value first (`writableCopy`,
 * `cloneDocumentWithParagraphPropertySources`).
 */
export type WritableRecord = {
  readonly [writableRecord]?: unknown[];
};

/** A model record's own fields, without the {@link WritableRecord} mark. */
export type RecordField<Value> = Exclude<keyof Value, keyof WritableRecord>;
