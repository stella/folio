/**
 * What a `w:tblLook` actually asks for.
 *
 * `CT_TblLook` states the same six facts twice: once as `w:val`, an
 * `ST_ShortHexNumber` bitmask, and once as an attribute per flag. `w:val` is
 * the older spelling and is all a pre-ECMA-376-2nd-edition producer writes, so
 * a reader that looks only at the flags renders such a table with no
 * conditional formatting at all.
 *
 * The model keeps both, exactly as authored, because absence, an explicit `0`
 * and an explicit `1` are three different documents. Resolving them into the
 * one answer a consumer wants — "is first-row formatting on?" — happens here
 * and nowhere else.
 */

import type { ExhaustiveFields, TableLook } from "../types/document";

/**
 * Every flag, in the order `CT_TblLook` declares them.
 *
 * The reader and the serializer both walk this, so the attributes go back out
 * in the order they came in.
 */
export const TABLE_LOOK_FLAGS = [
  "firstRow",
  "lastRow",
  "firstColumn",
  "lastColumn",
  "noHBand",
  "noVBand",
] as const;

export type TableLookFlag = (typeof TABLE_LOOK_FLAGS)[number];

/** The bit of `w:val` each flag restates (ECMA-376 §17.4.57). */
export const TABLE_LOOK_BITS = {
  firstRow: 0x00_20,
  lastRow: 0x00_40,
  firstColumn: 0x00_80,
  lastColumn: 0x01_00,
  noHBand: 0x02_00,
  noVBand: 0x04_00,
} as const satisfies Record<TableLookFlag, number>;

/**
 * Every flag, decided.
 *
 * Binding the record to `TableLook`'s own keys is the totality gate: a flag
 * added to the model that {@link TABLE_LOOK_FLAGS} does not list makes this
 * alias fail its constraint, so the module stops compiling rather than letting
 * a flag ship with no bit, no read order and no resolution.
 */
export type ResolvedTableLook = ExhaustiveFields<
  Record<TableLookFlag, boolean>,
  Exclude<keyof TableLook, "val">
>;

const tableLookMask = (val: string | undefined): number => {
  if (val === undefined) {
    return 0;
  }
  const bits = Number.parseInt(val, 16);
  // A `w:val` outside `ST_ShortHexNumber` states nothing, which leaves every
  // flag on its own default rather than inventing bits.
  return Number.isNaN(bits) ? 0 : bits;
};

/**
 * Resolve a `w:tblLook` to the six answers a consumer can act on.
 *
 * Precedence is the flag the author stated, then the matching bit of `w:val`,
 * then off. The attribute form wins because it is the later spelling of the
 * same fact: a producer that writes both and disagrees means the flags, and a
 * producer that writes only `w:val` still gets its banding and its header row.
 */
export const resolveTableLook = (look: TableLook | undefined): ResolvedTableLook => {
  const mask = tableLookMask(look?.val);
  const resolve = (flag: TableLookFlag): boolean =>
    // oxlint-disable-next-line no-bitwise -- w:val is an OOXML bitmask
    look?.[flag] ?? (mask & TABLE_LOOK_BITS[flag]) !== 0;

  return {
    firstRow: resolve("firstRow"),
    lastRow: resolve("lastRow"),
    firstColumn: resolve("firstColumn"),
    lastColumn: resolve("lastColumn"),
    noHBand: resolve("noHBand"),
    noVBand: resolve("noVBand"),
  };
};

type SetTableLookFlagsOptions = {
  look: TableLook | undefined;
  /**
   * The flags to state. A flag left out keeps whatever the author said, which
   * may be nothing; this is a patch over the authored value, not a decision
   * per flag, so `Partial` is the honest type.
   */
  flags: Readonly<Partial<ResolvedTableLook>>;
};

/**
 * State flags on a `w:tblLook`, leaving `w:val` as the author wrote it.
 *
 * The attribute form is what {@link resolveTableLook} reads first, so stating a
 * flag is enough to change the table. Re-encoding `w:val` instead would rewrite
 * the bits outside the six modelled here — the bitmask is 16 bits wide and the
 * format reserves the rest — and folio has nothing to put there.
 */
export const setTableLookFlags = ({ look, flags }: SetTableLookFlagsOptions): TableLook => ({
  ...look,
  ...flags,
});
