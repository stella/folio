// Fixture for `folio-at-wraparound/no-computed-negative-index`. Literal
// indexes, forward offsets and the non-wrapping helpers are the accepted
// spellings.

import { itemBefore } from "../../packages/core/src/utils/listItems";

export const last = (list: readonly string[]) => list.at(-1);

export const first = (list: readonly string[]) => list.at(0);

export const next = (list: readonly string[], index: number) =>
  list.at(index + 1);

export const previous = (list: readonly string[], index: number) =>
  itemBefore(list, index);

export const kept = (list: readonly string[], step: number) =>
  // oxlint-disable-next-line folio-at-wraparound/no-computed-negative-index -- the fixture's cyclic read
  list.at(step - 1);

export const plain = (list: readonly string[], index: number) =>
  list[index - 1];
