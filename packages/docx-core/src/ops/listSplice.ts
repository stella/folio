// A splice carries the identity of its immutable source without retaining it.
type SpliceBounds = { start: number; previousEnd: number; end: number };
const listKeys = new WeakMap<readonly unknown[], object>();
export const listKey = (list: readonly unknown[]): object => {
  let key = listKeys.get(list);
  if (key === undefined) {
    key = {};
    listKeys.set(list, key);
  }
  return key;
};

const splices = new WeakMap<
  readonly unknown[],
  SpliceBounds & { previous: object }
>();

type SpliceItemsOptions<T> = {
  items: readonly T[];
  start: number;
  count: number;
  replacement: readonly T[];
};

export const spliceItems = <T>({
  items,
  start,
  count,
  replacement,
}: SpliceItemsOptions<T>): T[] => {
  const next = items.toSpliced(start, count, ...replacement);
  if (
    Number.isInteger(start) &&
    start >= 0 &&
    start <= items.length &&
    Number.isInteger(count) &&
    count >= 0 &&
    count <= items.length - start
  ) {
    splices.set(next, {
      previous: listKey(items),
      start,
      previousEnd: start + count,
      end: start + replacement.length,
    });
  }
  return next;
};

export const sourceSplice = (current: readonly unknown[]) =>
  splices.get(current);

/** Outside these bounds both arrays contain precisely the same values. */
export const changedListBounds = (
  previous: readonly unknown[],
  current: readonly unknown[],
): SpliceBounds => {
  const splice = sourceSplice(current);
  if (splice !== undefined && splice.previous === listKeys.get(previous))
    return splice;
  // Inverse validation compares the same two versions in reverse order.
  const inverse = sourceSplice(previous);
  if (inverse !== undefined && inverse.previous === listKeys.get(current))
    return {
      start: inverse.start,
      previousEnd: inverse.end,
      end: inverse.previousEnd,
    };
  let start = 0;
  let previousEnd = previous.length;
  let end = current.length;
  while (
    start < previousEnd &&
    start < end &&
    previous[start] === current[start]
  )
    start++;
  while (
    previousEnd > start &&
    end > start &&
    previous[previousEnd - 1] === current[end - 1]
  ) {
    previousEnd--;
    end--;
  }
  return { start, previousEnd, end };
};
