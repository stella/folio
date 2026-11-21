// Fixture for `folio-at-wraparound/no-computed-negative-index`. The rule must
// flag every `.at()` read whose index is computed by subtraction or negation.

export const previous = (list: readonly string[], index: number) =>
  list.at(index - 1);

export const span = (list: readonly string[], end: number, start: number) =>
  list.at(end - start);

export const parenthesized = (list: readonly string[], index: number) =>
  list.at((index - 2)!);

export const fromEnd = (list: readonly string[], count: number) =>
  list.at(-count);

export const chained = (rows: readonly (readonly string[])[], row: number) =>
  rows
    .at(0)
    ?.at(row - 1);

export const text = (line: string, end: number) => line.at(end - 1);
