// Fixture for `folio-math-extremes/no-spread-arguments`. The rule must flag
// every way an array becomes the argument list of `Math.max` or `Math.min`.

export const largest = (values: readonly number[]): number => Math.max(...values);

export const lowest = (values: readonly number[]): number => Math.min(0, ...values);

export const viaGlobalThis = (values: readonly number[]): number =>
  globalThis.Math["max"](1, ...values);

export const viaApply = (values: readonly number[]): number =>
  Math.max.apply(null, [...values]);

export const viaReflect = (values: readonly number[]): number =>
  Reflect.apply(Math.min, undefined, values);
