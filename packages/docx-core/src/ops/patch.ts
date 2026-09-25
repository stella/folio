import type { FormattingPatch } from "./types";

/**
 * The keys an object states.
 *
 * SAFETY: `Object.keys` widens to `string[]`. Both callers pass a model
 * property set or a patch over one, object literals whose own keys are all
 * keys of `Formatting`.
 */
const ownKeys = <Formatting>(value: object): (keyof Formatting)[] =>
  Object.keys(value) as (keyof Formatting)[];

/**
 * A property set with a patch applied: a value sets its key, `null` clears
 * it, an absent key is untouched. The input is not modified. A set left with
 * no keys is `undefined`, the model's spelling of "states nothing".
 */
export const applyFormattingPatch = <Formatting extends object>(
  base: Formatting | undefined,
  patch: FormattingPatch<Formatting>,
): Partial<Formatting> | undefined => {
  const next: Partial<Formatting> = {};
  if (base !== undefined) {
    for (const key of ownKeys<Formatting>(base)) {
      if (patch[key] === undefined) {
        next[key] = base[key];
      }
    }
  }
  for (const key of ownKeys<Formatting>(patch)) {
    const value = patch[key];
    if (value !== undefined && value !== null) {
      next[key] = value;
    }
  }
  return Object.keys(next).length === 0 ? undefined : next;
};
