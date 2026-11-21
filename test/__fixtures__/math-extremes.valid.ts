// Fixture for `folio-math-extremes/no-spread-arguments`. Fixed argument lists
// and the folding helpers are the accepted spellings.

import { maxOf, minOf } from "../../packages/core/src/utils/extremes";

export const clamp = (value: number, low: number, high: number): number =>
  Math.min(high, Math.max(low, value));

export const largest = (values: readonly number[]): number => maxOf(values, 0);

export const lowest = (values: readonly number[]): number => minOf(values);

export const copy = (values: readonly number[]): number[] => [...values];
