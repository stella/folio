// Fixture for `folio-typecheck-proofs/no-type-suppression-in-test`. The rule
// must flag both directives, in a line comment and in a block comment.

export const widened: number =
  // @ts-expect-error a string is not a number
  "1";

export const ignored: number =
  /* @ts-ignore as above, spelled the other way */
  "2";

export const alsoWidened: number =
  // @ts-expect-error and a second one on its own line
  "3";
