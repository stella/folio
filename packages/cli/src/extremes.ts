// `Math.max(...values)` passes every element as a call argument, and engines
// throw `RangeError` once an array outgrows their argument limit. This folds
// instead and returns what the spread call returns: `maxOf(values, floor)` is
// `Math.max(floor, ...values)`, so an empty input yields the floor and a `NaN`
// element yields `NaN`. `.oxlint-plugins/folio-math-extremes.ts` keeps spread
// calls out of package source.

/** `Math.max(floor, ...values)` without the argument limit. */
export const maxOf = (
  values: Iterable<number>,
  floor = Number.NEGATIVE_INFINITY,
): number => {
  let result = floor;
  for (const value of values) {
    result = Math.max(result, value);
  }
  return result;
};
