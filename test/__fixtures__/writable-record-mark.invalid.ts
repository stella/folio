// Fixture for `folio-model-types/writable-record-mark`: every record below is
// declared without the `WritableRecord` mark.
declare const mark: unique symbol;
type WritableRecord = { readonly [mark]?: unknown[] };

export type Named = { a: number };
export type Outer = WritableRecord & { inner: { b: number } };
export type Fonts = WritableRecord & { fonts: Record<string, string> };
export type Picked = WritableRecord & { picked: Pick<Named, "a"> };
export type Listed = WritableRecord & { items: { c: number }[] };
export interface Plain {
  d: number;
}
export type Kinds = { kind: "a"; e: number } | { kind: "b" };
