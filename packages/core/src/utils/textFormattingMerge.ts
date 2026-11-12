import type { ColorValue, TextFormatting } from "../types/document";
import type { DeepReadonly, OwnRecord } from "../types/documentModel";
import { mergeFontFamily } from "./fontFamilyMerge";

// Double strike is deliberately absent: it uses ordinary last-defined inheritance.
export const STYLE_TOGGLE_KEYS: readonly [
  "bold",
  "boldCs",
  "italic",
  "italicCs",
  "allCaps",
  "emboss",
  "imprint",
  "outline",
  "shadow",
  "smallCaps",
  "strike",
  "hidden",
] = [
  "bold",
  "boldCs",
  "italic",
  "italicCs",
  "allCaps",
  "emboss",
  "imprint",
  "outline",
  "shadow",
  "smallCaps",
  "strike",
  "hidden",
] as const satisfies readonly (keyof TextFormatting)[];

/**
 * Merge one style's own run-formatting declarations onto its already-resolved
 * `w:basedOn` ancestor.
 *
 * ECMA-376 §17.7.3 resolves a toggle property (bold, italic, allCaps, ...) for a
 * single style tier by walking the `basedOn` chain from the style itself towards
 * its root: read the value the style declares; if it declares none, move to the
 * parent and repeat. The first value found — whether an explicit "on" or an
 * explicit "off" — is that tier's value, and it is *not* combined with whatever
 * a weaker ancestor further up the chain happened to say. A style's own
 * `<w:b w:val="0"/>` therefore turns bold off outright, even when its base style
 * is bold; it does not defer to the base.
 *
 * That is exactly last-defined-wins, i.e. `mergeTextFormatting`'s existing
 * behaviour for every other property. The toggle behaviour §17.7.3 is actually
 * known for — combining a table style's, a paragraph style's and a character
 * style's resolved values by Boolean XOR — combines the *outputs* of separate
 * `basedOn` walks like this one; it is not part of resolving a single walk, so
 * it has no place in this function.
 */
export function mergeStyleTextFormatting(
  target: DeepReadonly<TextFormatting> | undefined,
  source: DeepReadonly<TextFormatting> | undefined,
): DeepReadonly<TextFormatting> | undefined;
export function mergeStyleTextFormatting(
  target: DeepReadonly<TextFormatting> | undefined,
  source: DeepReadonly<TextFormatting> | undefined,
): DeepReadonly<TextFormatting> | undefined {
  return mergeTextFormatting(target, source);
}

/**
 * A merge answers "what does this text look like once the style, the paragraph
 * mark and the run have all had their say". `preserved` is not a value: it is
 * the bytes one specific `w:rPr` held, and carrying it across a merge would
 * write a style's or a paragraph mark's markup into every run that inherits
 * from it. So no merged formatting has one, on any path out of the function.
 */
const withoutPreserved = (
  formatting: DeepReadonly<TextFormatting>,
): DeepReadonly<TextFormatting> => {
  if (formatting.preserved === undefined) {
    return formatting;
  }
  const { preserved: _preserved, ...values } = formatting;
  return values;
};

/**
 * A merge's own record: a merge with a source always builds a new top-level
 * record, so its fields are the caller's to set, while the nested values in
 * it (an underline, a colour, a font family) stay shared with the inputs.
 */
export type MergedTextFormatting = OwnRecord<TextFormatting>;

/**
 * The merge shares nested values with its inputs and, without a source,
 * returns the target itself, so its result is read-only. A merge with a
 * source returns a new record ({@link MergedTextFormatting}) whose own fields
 * the caller may set. There is no writable variant: TypeScript lets a
 * readonly literal match a writable parameter, which would hand its shared
 * values out writable.
 */
export function mergeTextFormatting(
  target: DeepReadonly<TextFormatting> | undefined,
  source: DeepReadonly<TextFormatting>,
): MergedTextFormatting;
export function mergeTextFormatting(
  target: DeepReadonly<TextFormatting> | undefined,
  source: DeepReadonly<TextFormatting> | undefined,
): DeepReadonly<TextFormatting> | undefined;
export function mergeTextFormatting(
  target: DeepReadonly<TextFormatting> | undefined,
  source: DeepReadonly<TextFormatting> | undefined,
): DeepReadonly<TextFormatting> | undefined {
  if (!source && !target) {
    return undefined;
  }
  if (!source) {
    return target && withoutPreserved(target);
  }
  if (!target) {
    return withoutPreserved({ ...source });
  }

  const result: Record<string, unknown> = { ...withoutPreserved(target) };

  for (const key of Object.keys(source) as (keyof TextFormatting)[]) {
    const value = source[key];
    if (value === undefined || key === "preserved") {
      continue;
    }

    if (key === "fontFamily" && typeof value === "object") {
      result["fontFamily"] = mergeFontFamily(
        target.fontFamily,
        value as NonNullable<TextFormatting["fontFamily"]>,
      );
      continue;
    }

    if (key === "color" && typeof value === "object") {
      result["color"] = value as ColorValue;
      continue;
    }

    if (typeof value === "object" && !Array.isArray(value)) {
      result[key] = {
        ...(target[key] as Record<string, unknown> | undefined),
        ...(value as Record<string, unknown>),
      };
      continue;
    }

    result[key] = value;
  }

  return result as DeepReadonly<TextFormatting>;
}
