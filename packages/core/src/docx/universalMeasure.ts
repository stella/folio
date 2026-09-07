/**
 * ECMA-376's unit-carrying length, and the numbers Transitional counts it in.
 *
 * `ST_UniversalMeasure` spells a length as a decimal with its unit attached
 * (`155.85pt`, `2.54cm`). Strict producers write it wherever the type allows;
 * Transitional attributes are read as plain numbers in a fixed unit, so both
 * the verbatim-capture conversion and the typed projection have to agree on
 * what one of those strings is worth. They share this module so they cannot.
 */

/** The unit a Transitional attribute's number is counted in. */
export type MeasureUnit = "emu" | "halfPoints" | "hundredthPoints" | "twips";

/** `-?123.45pt`: the one shape ECMA-376 gives a length that carries its unit. */
const UNIVERSAL_MEASURE = /^(-?[0-9]+(?:\.[0-9]+)?)(mm|cm|in|pt|pc|pi)$/u;

const POINTS_PER_UNIT: Readonly<Record<string, number>> = {
  cm: 72 / 2.54,
  in: 72,
  mm: 72 / 25.4,
  pc: 12,
  pi: 12,
  pt: 1,
};

const NUMBERS_PER_POINT: Readonly<Record<MeasureUnit, number>> = {
  emu: 12_700,
  halfPoints: 2,
  hundredthPoints: 100,
  twips: 20,
};

/** Round half away from zero, so a negative offset keeps a positive one's magnitude. */
export const roundHalfAwayFromZero = (value: number): number =>
  value < 0 ? -Math.round(-value) : Math.round(value);

/** The value of a universal measure in points, or `undefined` if it is not one. */
export const universalMeasureInPoints = (value: string): number | undefined => {
  const match = UNIVERSAL_MEASURE.exec(value);
  if (match === null) {
    return undefined;
  }
  // SAFETY: both capture groups are present whenever the pattern matched.
  return Number(match[1]!) * POINTS_PER_UNIT[match[2]!]!;
};

/** A universal measure as the whole number a Transitional attribute of `unit` holds. */
export const universalMeasureAs = (value: string, unit: MeasureUnit): number | undefined => {
  const points = universalMeasureInPoints(value);
  return points === undefined ? undefined : roundHalfAwayFromZero(points * NUMBERS_PER_POINT[unit]);
};
