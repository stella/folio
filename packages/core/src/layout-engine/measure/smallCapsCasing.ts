/**
 * Casing rules for synthesized `w:smallCaps` glyphs.
 *
 * A face that carries no `smcp`/`c2sc` substitution has no true small-capital
 * glyph, so it is faked: a lowercase letter is drawn as its uppercase form at
 * a smaller point size, on the run's own baseline, at a fixed fraction of the
 * run's declared size — 0.8, an exact ratio observed to repeat across mixed
 * two-size runs at every point size and never round to a half-point `w:sz`
 * step. That is not a canvas or DOM backend's own `font-variant: small-caps`,
 * which synthesizes at its own, uncontrollable, narrower ratio (Blink and
 * WebKit use 0.7).
 *
 * Pure: no canvas, no DOM. Shared by every measurement and paint backend so a
 * run's advances and its glyphs agree on which characters shrink.
 */

/** Fraction of the run's declared size a synthesized small capital draws at. */
export const SMALL_CAPS_SCALE = 0.8;

/** Whitespace never gets its own small-caps class; see `smallCapsMask`. */
const WHITESPACE_PATTERN = /\s/u;

/**
 * Whether `char` is a cased letter in its lowercase form — the class
 * `w:smallCaps` draws as a shrunken, uppercased capital. Digits, punctuation,
 * symbols and already-uppercase letters keep the run's full size; combining
 * `char`'s upper and lower forms this way (rather than comparing to a
 * separate "is uppercase" test) is what correctly leaves caseless scripts
 * (CJK, most punctuation) at full size too, since neither form differs from
 * the character itself there.
 */
export function isSmallCapsLetter(char: string): boolean {
  return char.toLocaleUpperCase() !== char;
}

/**
 * Per-code-point small-caps classification for `text`, one entry per code
 * point (iterated the same way `[...text]`/`for...of` does, so it lines up
 * with any other code-point walk of the same string).
 *
 * Whitespace carries no case of its own; a run's space draws at the size of
 * the character immediately before it — the trailing space after a
 * synthesized word is itself synthesized — so a whitespace code point here
 * takes the previous code point's class, and only falls back to full size
 * when it opens the string (nothing to inherit from).
 */
export function smallCapsMask(text: string): boolean[] {
  const mask: boolean[] = [];
  let previous = false;
  for (const char of text) {
    const small: boolean = WHITESPACE_PATTERN.test(char) ? previous : isSmallCapsLetter(char);
    mask.push(small);
    previous = small;
  }
  return mask;
}

/** A maximal run of `text` drawn at one size: `small` selects which. */
export type SmallCapsSegment = {
  readonly text: string;
  readonly small: boolean;
};

/**
 * `text` split into maximal same-size segments for painting, each already
 * uppercased where `small` — the shrunken capital a `smcp`-less face
 * substitutes for the source lowercase letter. A run with no lowercase letter
 * at all (already-uppercase headings, digits, punctuation) comes back as one
 * `small: false` segment, so a caller can skip the extra painting work its
 * mixed-size runs need.
 */
export function smallCapsSegments(text: string): SmallCapsSegment[] {
  const mask = smallCapsMask(text);
  const segments: SmallCapsSegment[] = [];
  let current = "";
  let currentSmall: boolean | undefined;
  let index = 0;
  for (const char of text) {
    const small = mask[index] ?? false;
    if (currentSmall !== undefined && small !== currentSmall) {
      segments.push({ text: current, small: currentSmall });
      current = "";
    }
    current += small ? char.toLocaleUpperCase() : char;
    currentSmall = small;
    index += 1;
  }
  if (current) {
    segments.push({ text: current, small: currentSmall ?? false });
  }
  return segments;
}
