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

// A series of one-for-one edits at the same slot shares every other value
// with its first source. Keep that source's key, without retaining any lists.
const sameSlotOrigin = new WeakMap<readonly unknown[], object>();

type SpliceItemsOptions<T> = {
  items: readonly T[];
  start: number;
  count: number;
  replacement: readonly T[];
};

// Keep small splices on the native path without relying on the runtime's
// argument limit for large replacement windows.
const MAX_NATIVE_SPLICE_INSERTIONS = 1_024;

const spliceLargeReplacement = <T>({
  items,
  start,
  count,
  replacement,
}: SpliceItemsOptions<T>): T[] => {
  if (start === 0 && count === items.length) return Array.from(replacement);

  // Native toSpliced still owns deletion and returns a dense list of the
  // retained entries. Make room for replacements without another list.
  const next = items.toSpliced(start, count);
  const startInteger = Math.trunc(start) || 0;
  const at =
    startInteger < 0
      ? Math.max(0, items.length + startInteger)
      : Math.min(items.length, startInteger);
  const retained = next.length;
  next.length += replacement.length;
  next.copyWithin(at + replacement.length, at, retained);
  let index = at;
  for (const item of replacement) next[index++] = item;
  return next;
};

export const spliceItems = <T>({
  items,
  start,
  count,
  replacement,
}: SpliceItemsOptions<T>): T[] => {
  const next =
    replacement.length <= MAX_NATIVE_SPLICE_INSERTIONS
      ? items.toSpliced(start, count, ...replacement)
      : spliceLargeReplacement({ items, start, count, replacement });
  if (
    Number.isInteger(start) &&
    start >= 0 &&
    start <= items.length &&
    Number.isInteger(count) &&
    count >= 0 &&
    count <= items.length - start
  ) {
    const previous = listKey(items);
    if (count === 1 && replacement.length === 1) {
      const prior = splices.get(items);
      const origin =
        prior?.start === start ? sameSlotOrigin.get(items) : undefined;
      sameSlotOrigin.set(next, origin ?? previous);
    }
    splices.set(next, {
      previous,
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
  const laterOrigin = sameSlotOrigin.get(current);
  const earlierOrigin = sameSlotOrigin.get(previous);
  // Planning and replay may independently edit the same slot from the same
  // source. Replacing another slot or shifting any position resets the proof.
  if (laterOrigin !== undefined && splice !== undefined) {
    if (
      laterOrigin === listKeys.get(previous) ||
      (laterOrigin === earlierOrigin && splice.start === inverse?.start)
    )
      return { start: splice.start, previousEnd: splice.end, end: splice.end };
  }
  if (
    earlierOrigin !== undefined &&
    inverse !== undefined &&
    earlierOrigin === listKeys.get(current)
  )
    return { start: inverse.start, previousEnd: inverse.end, end: inverse.end };
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
