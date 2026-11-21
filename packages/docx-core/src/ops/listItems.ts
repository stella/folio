/**
 * The item of `list` just before `index`; `undefined` at the list's start,
 * where `list.at(index - 1)` would read the list's last item instead.
 */
export const itemBefore = <T>(
  list: readonly T[],
  index: number,
): T | undefined => (index > 0 ? list[index - 1] : undefined);
