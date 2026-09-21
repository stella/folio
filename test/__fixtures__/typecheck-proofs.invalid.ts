// @ts-nocheck
// Fixture for `folio-typecheck-proofs/no-type-suppression-in-test`. The rule
// must flag all suppression directives, in line and block comments.

export const widened: number =
  // @ts-expect-error a string is not a number
  "1";

export const ignored: number =
  /* @ts-ignore as above, spelled the other way */
  "2";

export const alsoWidened: number =
  // @ts-expect-error and a second one on its own line
  "3";
