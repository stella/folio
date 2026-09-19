/**
 * The reserved `w:numId` value that means "no numbering".
 *
 * ECMA-376 (17.9.18 `numId`, 17.9.19 `numPr`) reserves 0: it names no `w:num`,
 * it switches numbering off, and on a style it cancels the numbering the style
 * would otherwise inherit through `w:basedOn` (Word writes it on, for example,
 * a heading-based style that must not be numbered). It must therefore survive a
 * round trip: dropping the `w:numPr` would hand the numbering back.
 */
export const NO_NUMBERING_NUM_ID = 0;

/**
 * Whether a `w:numId` names a numbering definition to resolve.
 *
 * Every lookup or validation against `numbering.nums` goes through this: an
 * absent id has nothing to resolve, and {@link NO_NUMBERING_NUM_ID} is the
 * "none" sentinel, never a dangling reference.
 */
export const isNumberingReference = (numId: number | undefined): numId is number =>
  numId !== undefined && numId !== NO_NUMBERING_NUM_ID;
