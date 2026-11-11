/**
 * A measure holds no document positions.
 *
 * A block's measure is reused across layout passes for a block that an edit
 * only moved: its lines, widths and heights stay right wherever the block
 * now starts. That holds only while no measure type records where a block or
 * run stands in the document. This module makes it a compile-time fact: a
 * measure type that gains a position-like field (`pmStart`, `fromPos`, …)
 * fails to compile here, before reuse can hand out a stale position.
 */

import type { Measure } from "../types";

/** Field names that say where something stands in the document. */
type PositionLikeKey<Key> = Key extends
  `pm${string}` | `${string}Pos` | `${string}Position` | "pos"
  ? Key
  : never;

/** Nesting deep enough for every measure type; measures nest measures. */
type MaxDepth = 8;

/** Every position-like field name found anywhere inside `Value`. */
export type PositionKeysIn<
  Value,
  Depth extends readonly unknown[] = [],
> = Depth["length"] extends MaxDepth
  ? never
  : Value extends readonly (infer Element)[]
    ? PositionKeysIn<Element, [...Depth, unknown]>
    : Value extends object
      ? {
          [Key in keyof Value]-?:
            | (Key extends string ? PositionLikeKey<Key> : never)
            | PositionKeysIn<NonNullable<Value[Key]>, [...Depth, unknown]>;
        }[keyof Value]
      : never;

/** Position-like field names a measure type holds: none. */
export type MeasurePositionKeys = PositionKeysIn<Measure>;

/**
 * True, and compiles only while {@link MeasurePositionKeys} is empty. A
 * measure type with a position field turns this declaration into a type
 * error that names the field.
 */
export const MEASURES_HOLD_NO_POSITIONS: [MeasurePositionKeys] extends [never]
  ? true
  : { positionFields: MeasurePositionKeys } = true;
