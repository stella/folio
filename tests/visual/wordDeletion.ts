/**
 * Deletion by word as the platform convention measures it, the browser naming
 * the range in its beforeinput: backward to the previous word's start,
 * forward to the next word's end, the whitespace on the way going with the
 * word. A character that is neither whitespace nor part of a word is a word
 * of its own. The one word-boundary oracle of the browser input replays.
 */
export type WordDeletionDirection = "backward" | "forward";

const isWhitespace = (character: string) => /\s/u.test(character);
const isWordCharacter = (character: string) =>
  /[\p{L}\p{M}\p{N}_]/u.test(character);

/** One word deletion from `caret` in `text`: what remains, the caret after it and what it deleted. */
export const wordDeletion = (
  text: string,
  caret: number,
  direction: WordDeletionDirection,
): { text: string; caret: number; deleted: string } => {
  const step = direction === "backward" ? -1 : 1;
  const inside = (edge: number) =>
    direction === "backward" ? edge > 0 : edge < text.length;
  const next = (edge: number) =>
    text.charAt(direction === "backward" ? edge - 1 : edge);
  let edge = caret;
  while (inside(edge) && isWhitespace(next(edge))) edge += step;
  if (inside(edge) && isWordCharacter(next(edge)))
    while (inside(edge) && isWordCharacter(next(edge))) edge += step;
  else if (inside(edge)) edge += step;
  const from = Math.min(caret, edge);
  const to = Math.max(caret, edge);
  return {
    text: text.slice(0, from) + text.slice(to),
    caret: from,
    deleted: text.slice(from, to),
  };
};

/**
 * Successive word deletions from `caret`, each from where the last one left
 * the caret: what remains, and the one span of `text` they removed together.
 */
export const wordDeletions = (
  text: string,
  caret: number,
  directions: readonly WordDeletionDirection[],
): { text: string; removed: string } => {
  let state = { text, caret };
  for (const direction of directions)
    state = wordDeletion(state.text, state.caret, direction);
  return {
    text: state.text,
    removed: text.slice(
      state.caret,
      state.caret + text.length - state.text.length,
    ),
  };
};
