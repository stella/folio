/**
 * Linear scans over the end of a text run. Layout calls these on every
 * measure and paint pass, and run text can be arbitrarily long, so they
 * must not backtrack the way an end-anchored regular expression does.
 */

const WHITESPACE = /\s/u;

const isWhitespaceUnit = (text: string, index: number): boolean =>
  WHITESPACE.test(text.charAt(index));

/** Index where the trailing run of U+0020 spaces begins; `text.length` when there is none. */
export function trailingSpaceStart(text: string): number {
  let start = text.length;
  while (start > 0 && text.charCodeAt(start - 1) === 0x20) {
    start--;
  }
  return start;
}

export type TrailingToken = {
  /** Last run of non-whitespace characters. */
  token: string;
  /** Whitespace after the token, up to the end of the text. */
  separator: string;
};

/**
 * Split off the last whitespace-delimited token and the whitespace that
 * follows it, or return undefined when the text has no non-whitespace
 * character.
 */
export function splitTrailingToken(text: string): TrailingToken | undefined {
  let tokenEnd = text.length;
  while (tokenEnd > 0 && isWhitespaceUnit(text, tokenEnd - 1)) {
    tokenEnd--;
  }
  if (tokenEnd === 0) {
    return undefined;
  }
  let tokenStart = tokenEnd;
  while (tokenStart > 0 && !isWhitespaceUnit(text, tokenStart - 1)) {
    tokenStart--;
  }
  return { token: text.slice(tokenStart, tokenEnd), separator: text.slice(tokenEnd) };
}
