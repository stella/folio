// Fixture for `folio-typecheck-proofs/no-type-suppression-in-test`. A test
// states its claims through types and assertions; prose about a directive is
// not a directive.

export const stated: number = 1;

/** A doc comment may name `@ts-expect-error` without being one. */
export const described = (value: number): string => String(value);

// Prose that mentions @ts-ignore mid-sentence suppresses nothing either.
export const alsoDescribed = (value: string): number => Number(value);
