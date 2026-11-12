// Fixture for `folio-typecheck-proofs/no-unchecked-field-gate`: a table keyed
// by a local value's own keys, and a plain shape check, are not field gates.
type Model = { a: number; b: string };
const LOCAL = { a: 1 } as const;

export const byValue = { a: "x" } satisfies Record<keyof typeof LOCAL, string>;
export const shaped = { a: 1, b: "" } satisfies Model;
