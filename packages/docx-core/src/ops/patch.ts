import type { FormattingPatch } from "./types";

/**
 * A property set with a patch applied: a value (including explicit undefined)
 * sets its key, null removes it, and an absent key is untouched. The input is
 * not modified. A set left with no keys is undefined, meaning "states nothing".
 */
export const applyFormattingPatch = <Formatting extends object>(
  base: Formatting | undefined,
  patch: FormattingPatch<Formatting>,
): Partial<Formatting> | undefined => {
  const next: Partial<Formatting> = base === undefined ? {} : { ...base };
  for (const [key, value] of Object.entries(patch)) {
    if (value === null) Reflect.deleteProperty(next, key);
    else Reflect.set(next, key, value);
  }
  return Object.keys(next).length === 0 ? undefined : next;
};

/**
 * The patch that gives back what `patch` replaced in `base`: each key it names
 * set to the value it had, or cleared where it had none.
 */
export const priorValues = <Formatting extends object>(
  base: Formatting | undefined,
  patch: FormattingPatch<Formatting>,
): FormattingPatch<Formatting> => {
  const restore: FormattingPatch<Formatting> = {};
  for (const key of Object.keys(patch)) {
    Reflect.set(
      restore,
      key,
      base !== undefined && Object.hasOwn(base, key) ? Reflect.get(base, key) : null,
    );
  }
  return restore;
};
