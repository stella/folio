/**
 * Where an offset into a block's clean text may fall.
 *
 * Offsets count UTF-16 code units, as JavaScript strings do. One character a
 * reader sees can span several: a character outside the Basic Multilingual
 * Plane is two (a surrogate pair), and a grapheme cluster — a letter and its
 * combining marks, an emoji joined with zero-width joiners, a flag's two
 * regional indicators — is several code points. An editing caret stands
 * only between two clusters.
 *
 * An offset between the two halves of a surrogate pair names no character at
 * all, and text cut there cannot be written: each half becomes a lone
 * surrogate, which XML cannot hold, so the save drops both and the character
 * disappears. Every offset an operation takes is refused there.
 *
 * An offset inside a grapheme cluster names a code point, and marks may start
 * there — a document can itself format a combining mark apart from its base,
 * and a comparison reproducing that formatting needs the same boundary — but
 * text written or broken there changes a character nobody named: the accent
 * of `é` moves onto the replacement, a family emoji falls apart into people.
 * Operations that change text or break a paragraph are refused there as well.
 */

const isHighSurrogate = (code: number): boolean => code >= 0xd8_00 && code <= 0xdb_ff;
const isLowSurrogate = (code: number): boolean => code >= 0xdc_00 && code <= 0xdf_ff;

/** Whether `offset` falls between the two halves of one surrogate pair of `text`. */
export const splitsSurrogatePair = (text: string, offset: number): boolean =>
  offset > 0 &&
  offset < text.length &&
  isHighSurrogate(text.charCodeAt(offset - 1)) &&
  isLowSurrogate(text.charCodeAt(offset));

let graphemeSegmenter: Intl.Segmenter | null | undefined;

const getGraphemeSegmenter = (): Intl.Segmenter | null => {
  if (graphemeSegmenter === undefined) {
    graphemeSegmenter =
      typeof Intl === "object" && typeof Intl.Segmenter === "function"
        ? new Intl.Segmenter(undefined, { granularity: "grapheme" })
        : null;
  }
  return graphemeSegmenter;
};

const MAX_CACHED_TEXTS = 16;
const graphemeStartsByText = new Map<string, ReadonlySet<number>>();

/**
 * The offsets at which a grapheme cluster of `text` starts. A batch asks about
 * the same block for each of its operations, so the last few answers are kept.
 */
const graphemeStarts = (text: string, segmenter: Intl.Segmenter): ReadonlySet<number> => {
  const cached = graphemeStartsByText.get(text);
  if (cached !== undefined) {
    return cached;
  }
  const starts = new Set<number>();
  for (const { index } of segmenter.segment(text)) {
    starts.add(index);
  }
  if (graphemeStartsByText.size >= MAX_CACHED_TEXTS) {
    const oldest = graphemeStartsByText.keys().next();
    if (oldest.done !== true) {
      graphemeStartsByText.delete(oldest.value);
    }
  }
  graphemeStartsByText.set(text, starts);
  return starts;
};

/**
 * Whether `offset` falls inside one grapheme cluster of `text`, or between
 * the halves of a surrogate pair. Without `Intl.Segmenter` only the surrogate
 * pair is recognised.
 */
export const splitsGraphemeCluster = (text: string, offset: number): boolean => {
  if (offset <= 0 || offset >= text.length) {
    return false;
  }
  if (splitsSurrogatePair(text, offset)) {
    return true;
  }
  const segmenter = getGraphemeSegmenter();
  return segmenter !== null && !graphemeStarts(text, segmenter).has(offset);
};

export type CharacterBoundaryStrictness =
  /** Refuse only an offset inside a surrogate pair: annotations. */
  | "codePoint"
  /** Refuse any offset inside a grapheme cluster: edits of text and breaks. */
  | "grapheme";

const splitsCharacter = (
  text: string,
  offset: number,
  strictness: CharacterBoundaryStrictness,
): boolean =>
  strictness === "grapheme"
    ? splitsGraphemeCluster(text, offset)
    : splitsSurrogatePair(text, offset);

/**
 * The character a split `offset` falls inside, as the offsets of its ends: the
 * whole grapheme cluster, or — without `Intl.Segmenter`, which leaves only a
 * surrogate pair recognisable — the pair.
 */
const enclosingCharacter = (text: string, offset: number): { start: number; end: number } => {
  const segmenter = getGraphemeSegmenter();
  if (segmenter === null) {
    return { start: offset - 1, end: offset + 1 };
  }
  const starts = graphemeStarts(text, segmenter);
  let start = offset;
  while (start > 0 && !starts.has(start)) {
    start--;
  }
  let end = offset;
  while (end < text.length && !starts.has(end)) {
    end++;
  }
  return { start, end };
};

/**
 * Why `offsets` do not all stand on a character boundary of `text`, naming
 * the first that does not and the character it cuts; `null` when they all do.
 */
export const describeCharacterSplit = (
  text: string,
  offsets: readonly number[],
  strictness: CharacterBoundaryStrictness,
): string | null => {
  const offset = offsets.find((candidate) => splitsCharacter(text, candidate, strictness));
  if (offset === undefined) {
    return null;
  }
  const { start, end } = enclosingCharacter(text, offset);
  return (
    `offset ${String(offset)} falls inside ${JSON.stringify(text.slice(start, end))}, ` +
    `which spans offsets ${String(start)} to ${String(end)}; use ${String(start)} or ${String(end)}.`
  );
};
