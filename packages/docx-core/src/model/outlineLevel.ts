import { panic } from "better-result";

/**
 * `w:outlineLvl`: nine heading levels and one reserved "not a heading".
 *
 * ECMA-376 17.3.1.20 gives `w:val` the range 0..9 "where 9 specifically
 * indicates that there is no outline level specifically applied to this
 * paragraph". Nine is therefore a sentinel, not a tenth level, and a model
 * that stores the raw number lets every consumer decide that for itself: the
 * PDF outline nested a `TOC Heading` nine deep, the TOC listed it, and the
 * lint could not see the mistake because the comparison was against
 * `undefined` rather than against the literal.
 *
 * The union removes the decision. `9` has no representation as a heading, an
 * out-of-range value has none at all, and absence stays absence — a paragraph
 * that states no level inherits one, which is a third state neither arm may
 * stand in for.
 */

/** The `w:outlineLvl w:val` values that name a heading, zero-based. */
const HEADING_OUTLINE_LEVELS = [0, 1, 2, 3, 4, 5, 6, 7, 8] as const;

/** A zero-based heading level: `heading 1` is 0, `heading 9` is 8. */
export type HeadingOutlineLevel = (typeof HEADING_OUTLINE_LEVELS)[number];

/** What a stated `w:outlineLvl` says. Absence is the field being absent. */
export type OutlineLevel =
  | { readonly kind: "bodyText" }
  | { readonly kind: "heading"; readonly level: HeadingOutlineLevel };

/** The reserved `w:outlineLvl w:val="9"`: outlined as body text, not as a heading. */
export const BODY_TEXT_OUTLINE_LEVEL: OutlineLevel = { kind: "bodyText" };

/** The `w:val` the body-text arm is written as. */
const BODY_TEXT_OUTLINE_LEVEL_VALUE = 9;

const OUTLINE_LEVEL_BY_STATED_VALUE: ReadonlyMap<number, OutlineLevel> = new Map<
  number,
  OutlineLevel
>([
  ...HEADING_OUTLINE_LEVELS.map((level): readonly [number, OutlineLevel] => [
    level,
    { kind: "heading", level },
  ]),
  [BODY_TEXT_OUTLINE_LEVEL_VALUE, BODY_TEXT_OUTLINE_LEVEL],
]);

/**
 * The one reader over a stated `w:outlineLvl w:val`.
 *
 * `undefined` is a value outside 0..9, which the format does not define. The
 * Rust kernel rejects those outright (`MAXIMUM_OUTLINE_LEVEL`); the TypeScript
 * parse boundary drops them and warns instead, because a corpus document that
 * Word opens must still open here.
 */
export const outlineLevelFromStatedValue = (value: number): OutlineLevel | undefined =>
  OUTLINE_LEVEL_BY_STATED_VALUE.get(value);

/** The heading arm for a zero-based level, or undefined outside 0..8. */
export const headingOutlineLevel = (level: number): OutlineLevel | undefined => {
  const outlineLevel = OUTLINE_LEVEL_BY_STATED_VALUE.get(level);
  return outlineLevel?.kind === "heading" ? outlineLevel : undefined;
};

/** The `w:val` to write for a stated outline level. */
export const outlineLevelStatedValue = (outlineLevel: OutlineLevel): number => {
  switch (outlineLevel.kind) {
    case "bodyText":
      return BODY_TEXT_OUTLINE_LEVEL_VALUE;
    case "heading":
      return outlineLevel.level;
    default: {
      const unhandled: never = outlineLevel;
      return panic(`Unhandled outline level ${JSON.stringify(unhandled)}`);
    }
  }
};

/**
 * The zero-based heading level a stated outline level names, or undefined when
 * it names none. Body text and an absent level answer the same thing, which is
 * what every heading consumer wants; the two are distinguished only where the
 * cascade still has a tier left to ask.
 */
export const headingLevelOf = (
  outlineLevel: OutlineLevel | null | undefined,
): HeadingOutlineLevel | undefined =>
  outlineLevel?.kind === "heading" ? outlineLevel.level : undefined;
