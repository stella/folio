/**
 * Text Selection Utilities
 *
 * Pure word-level boundary detection used for double-click (word) selection.
 * The DOM-driven selection + multi-click helpers live in the adapter layer.
 * @packageDocumentation
 * @public
 */

/**
 * Regular expression for word characters.
 * Includes letters, numbers, combining marks, underscores, and
 * common word-internal punctuation (apostrophes, hyphens).
 */
const WORD_CHAR_REGEX = /[\p{L}\p{N}\p{M}_''\-]/u;

/**
 * Regular expression for whitespace characters
 */
const WHITESPACE_REGEX = /\s/;

/**
 * Check if a character is a word character
 */
export function isWordCharacter(char: string): boolean {
  if (!char || char.length === 0) return false;
  return WORD_CHAR_REGEX.test(char);
}

/**
 * Check if a character is whitespace
 */
export function isWhitespace(char: string): boolean {
  if (!char || char.length === 0) return false;
  return WHITESPACE_REGEX.test(char);
}

/**
 * Find word boundaries around a position in text.
 * Returns [startIndex, endIndex] inclusive start, exclusive end.
 */
export function findWordBoundaries(text: string, position: number): [number, number] {
  if (!text || text.length === 0) {
    return [0, 0];
  }

  // Clamp position to valid range
  position = Math.max(0, Math.min(position, text.length - 1));

  const charAtPosition = text[position];

  // If on whitespace, select the whitespace run
  if (isWhitespace(charAtPosition)) {
    let start = position;
    let end = position;

    while (start > 0 && isWhitespace(text[start - 1])) {
      start--;
    }
    while (end < text.length && isWhitespace(text[end])) {
      end++;
    }

    return [start, end];
  }

  // If on a word character, find the word
  if (isWordCharacter(charAtPosition)) {
    let start = position;
    let end = position;

    while (start > 0 && isWordCharacter(text[start - 1])) {
      start--;
    }
    while (end < text.length && isWordCharacter(text[end])) {
      end++;
    }

    return [start, end];
  }

  // On punctuation or other non-word character, just select that character
  return [position, position + 1];
}

/**
 * Get the word at a position in text
 */
export function getWordAt(text: string, position: number): string {
  const [start, end] = findWordBoundaries(text, position);
  return text.slice(start, end);
}

/**
 * Word selection result
 */
export interface WordSelectionResult {
  /** The selected word */
  word: string;
  /** Start index in the text (inclusive) */
  startIndex: number;
  /** End index in the text (exclusive) */
  endIndex: number;
}

/**
 * Find the word at a position and return detailed info
 */
export function findWordAt(text: string, position: number): WordSelectionResult {
  const [start, end] = findWordBoundaries(text, position);
  return {
    word: text.slice(start, end),
    startIndex: start,
    endIndex: end,
  };
}
