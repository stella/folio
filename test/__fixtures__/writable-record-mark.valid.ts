// Fixture for `folio-model-types/writable-record-mark`: every record carries
// the mark, is read-only throughout, or is a phantom brand; a function's
// parameter type is not a model record.
declare const mark: unique symbol;
declare const brand: unique symbol;
type WritableRecord = { readonly [mark]?: unknown[] };

export type Named = WritableRecord & {
  a: number;
  inner: WritableRecord & { b: number };
};
export type Listed = WritableRecord & {
  items: (WritableRecord & { c: number })[];
};
export type Fonts = WritableRecord & {
  fonts: WritableRecord & Record<string, string>;
};
export type Picked = WritableRecord & Pick<Named, "a">;
export type Kinds = WritableRecord & ({ kind: "a" } | { kind: "b" });
export type Presence = WritableRecord & Partial<{ d: number }>;
export type Frozen = { readonly x: number; readonly y: number };
export type Branded = { readonly [brand]: true };
export interface Extended extends WritableRecord {
  e: number;
}
export const reader = (value: { f: number }): number => value.f;
