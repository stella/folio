// Fixture for `folio-typecheck-proofs/no-unchecked-field-gate`: each value is
// a gate total over a declared type's fields.
type Model = { a: number; b: string };
type RecordField<Value> = keyof Value;

export const total = { a: 1, b: "" } satisfies Required<Model>;
export const byKey = { a: "a", b: "b" } satisfies Record<keyof Model, string>;
export const byField = { a: 1, b: 2 } satisfies Record<
  RecordField<Model>,
  number
>;
export const mapped = { a: 1, b: "" } satisfies {
  [Field in keyof Model]-?: Model[Field];
};
