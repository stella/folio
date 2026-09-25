/**
 * The smallest set of text changes that turns one clean-text span into another.
 *
 * A direct-mode replacement (`replaceInBlock`, `replaceRange`, `replaceBlock`)
 * names the text it matched and the text it wants. Swapping the whole match
 * gives every character of the result the formatting of the first one, drops
 * the inline content controls, note references and fields the match crossed,
 * and makes a one-character edit rewrite the paragraph. The edit is the
 * difference between the two strings, so that is all the applier writes:
 *
 * - Text outside a change is not touched: its runs, marks and every inline node
 *   between its characters stay exactly as they were.
 * - A change deletes only the characters it removes (text, tabs, breaks and
 *   whole field results). Non-text inline content between those characters —
 *   bookmarks, comment anchors, drawings, content-control boundaries, rendered
 *   page breaks — stays where it was; the new text is written where the first
 *   removed character stood and takes its formatting.
 * - A pure insertion follows the character before it, with the formatting that
 *   character has, the way typed text does. At the start of the matched span
 *   there is no such character inside the match, so it takes the formatting of
 *   the first matched character. A note reference's number lends no
 *   formatting: prose beside it is not superscript.
 * - An insertion at the edge of an inline content control (`w:sdt`), a link or
 *   a comment range stays inside it only when the whole match lies inside it.
 *   Matching a control's or a link's text edits it; appending to a paragraph
 *   that ends in a checkbox or a link writes after it.
 * - A field result is one unit: a change that would cut into it is widened to
 *   the whole result, which the new text then replaces as plain text (a result
 *   cannot be partly edited; it is regenerated from the field code). A field
 *   whose displayed text the replacement keeps is never touched. A match that
 *   itself starts or ends inside a field result is refused as before
 *   (`unsupportedBlock`).
 *
 * A replacement carrying inline emphasis markup (`**bold**`) states its own
 * formatting and still replaces the match whole in direct mode.
 *
 * Tracked-changes and suggested modes cut their changes from the redline diff
 * instead ({@link changesFromSegments}: word or character granularity, as the
 * caller asked), and map them onto the document by the same rules. A change
 * marks exactly the characters it removes as a deletion, which keeps their
 * runs, and writes its text as an insertion after them, formatted as the
 * first removed character that is not a note reference; a pure insertion is
 * placed and formatted as above. Nothing outside a change is wrapped in a
 * revision.
 *
 * This module is the pure half: clean-text offsets in, changes out. The
 * applier maps each change onto the document.
 */

import { tokenizeWords, type WordDiffSegment } from "./word-diff";

/** Replace `source.slice(start, end)` with `text`; offsets index the source span. */
export type TextChange = { start: number; end: number; text: string };

/** A clean-text span that can only be replaced whole (a field result). */
export type AtomicTextSpan = { offset: number; length: number };

const isHighSurrogate = (code: number): boolean => code >= 0xd8_00 && code <= 0xdb_ff;
const isLowSurrogate = (code: number): boolean => code >= 0xdc_00 && code <= 0xdf_ff;

/** Common prefix length that never ends between the halves of a surrogate pair. */
export const commonPrefixLength = (a: string, b: string): number => {
  const limit = Math.min(a.length, b.length);
  let index = 0;
  while (index < limit && a.charCodeAt(index) === b.charCodeAt(index)) {
    index++;
  }
  // Stopping after a high surrogate would split its pair.
  return index > 0 && isHighSurrogate(a.charCodeAt(index - 1)) ? index - 1 : index;
};

/** Common suffix length that never starts between the halves of a surrogate pair. */
export const commonSuffixLength = (a: string, b: string): number => {
  const limit = Math.min(a.length, b.length);
  let length = 0;
  while (
    length < limit &&
    a.charCodeAt(a.length - 1 - length) === b.charCodeAt(b.length - 1 - length)
  ) {
    length++;
  }
  // Starting at a low surrogate would split its pair.
  return length > 0 && isLowSurrogate(a.charCodeAt(a.length - length)) ? length - 1 : length;
};

/** `change` without the characters its deleted and inserted text share at either end. */
const trimChange = (source: string, change: TextChange): TextChange | null => {
  const removed = source.slice(change.start, change.end);
  const prefix = commonPrefixLength(removed, change.text);
  const suffix = commonSuffixLength(removed.slice(prefix), change.text.slice(prefix));
  const trimmed = {
    start: change.start + prefix,
    end: change.end - suffix,
    text: change.text.slice(prefix, change.text.length - suffix),
  };
  return trimmed.start === trimmed.end && trimmed.text.length === 0 ? null : trimmed;
};

const reconstructs = (
  segments: readonly WordDiffSegment[],
  source: string,
  replacement: string,
): boolean => {
  let before = "";
  let after = "";
  for (const segment of segments) {
    if (segment.type !== "ins") {
      before += segment.text;
    }
    if (segment.type !== "del") {
      after += segment.text;
    }
  }
  return before === source && after === replacement;
};

/**
 * Edits beyond which a replacement is treated as one rewrite. Bounds the diff
 * at O((N + M) * D) time and O(D^2) memory; a rewrite that large keeps
 * nothing of its source worth aligning to anyway.
 */
export const MAX_TOKEN_EDITS = 1_024;

/**
 * The shortest edit script between two token sequences (Myers, 1986), as
 * segments of joined token text, or `null` past `maxEdits` edits.
 *
 * Deliberately the shortest one. The redline diff in `word-diff.ts` gives up
 * short matches between changes so a reader sees whole phrases replaced; an
 * edit applied directly has no reader to serve, and every token it gives up is
 * text rewritten with someone else's formatting.
 */
export const shortestTokenDiff = (
  before: readonly string[],
  after: readonly string[],
  maxEdits: number = MAX_TOKEN_EDITS,
): WordDiffSegment[] | null => {
  const n = before.length;
  const m = after.length;
  const limit = Math.min(n + m, maxEdits);
  const offset = limit + 1;
  const frontier = new Int32Array(2 * limit + 3);
  // `trace[d]` holds the frontier for diagonals -d..d before round d.
  const trace: Int32Array[] = [];
  let edits = -1;
  search: for (let d = 0; d <= limit; d++) {
    trace.push(frontier.slice(offset - d, offset + d + 1));
    for (let k = -d; k <= d; k += 2) {
      const down = k === -d || (k !== d && frontier[offset + k - 1]! < frontier[offset + k + 1]!);
      let x = down ? frontier[offset + k + 1]! : frontier[offset + k - 1]! + 1;
      let y = x - k;
      while (x < n && y < m && before[x] === after[y]) {
        x++;
        y++;
      }
      frontier[offset + k] = x;
      if (x >= n && y >= m) {
        edits = d;
        break search;
      }
    }
  }
  if (edits < 0) {
    return null;
  }

  const reversed: WordDiffSegment[] = [];
  const push = (type: WordDiffSegment["type"], text: string) => {
    const last = reversed.at(-1);
    if (last?.type === type) {
      last.text = text + last.text;
    } else {
      reversed.push({ type, text });
    }
  };
  let x = n;
  let y = m;
  for (let d = edits; d > 0; d--) {
    // SAFETY: one trace entry was pushed per round up to `edits`.
    const previous = trace[d]!;
    const at = (diagonal: number): number => previous[diagonal + d]!;
    const k = x - y;
    const down = k === -d || (k !== d && at(k - 1) < at(k + 1));
    const previousK = down ? k + 1 : k - 1;
    const previousX = at(previousK);
    const previousY = previousX - previousK;
    while (x > previousX && y > previousY) {
      x--;
      y--;
      push("equal", before[x]!);
    }
    if (down) {
      y--;
      push("ins", after[y]!);
    } else {
      x--;
      push("del", before[x]!);
    }
  }
  while (x > 0 && y > 0) {
    x--;
    y--;
    push("equal", before[x]!);
  }
  return reversed.toReversed();
};

/**
 * The changes from `source` to `replacement`, in source order and disjoint.
 *
 * Words (with their leading whitespace, and edge punctuation on its own; see
 * `tokenizeWords`) are aligned by the shortest edit script, and each changed
 * stretch is then trimmed to the characters that actually differ, so
 * appending a letter to a word that spans two runs leaves both runs alone. A
 * rewrite past {@link MAX_TOKEN_EDITS}, or a diff that does not reconstruct
 * both strings exactly (it never should), is one change over the whole span,
 * trimmed the same way.
 */
export const planTextChanges = (source: string, replacement: string): TextChange[] => {
  if (source === replacement) {
    return [];
  }
  const diffed = shortestTokenDiff(tokenizeWords(source), tokenizeWords(replacement));
  const segments: readonly WordDiffSegment[] =
    diffed !== null && reconstructs(diffed, source, replacement)
      ? diffed
      : [
          { type: "del", text: source },
          { type: "ins", text: replacement },
        ];

  return changesFromSegments(segments).flatMap((change) => trimChange(source, change) ?? []);
};

/**
 * The changes a diff's segments describe, in source order and disjoint: each
 * run of `del` and `ins` segments between two `equal` ones is one change,
 * exactly as wide as the segments say. A redline's segments come cut to the
 * granularity its reader asked for, so they are not trimmed further.
 */
export const changesFromSegments = (segments: readonly WordDiffSegment[]): TextChange[] => {
  const changes: TextChange[] = [];
  let cursor = 0;
  let pending: TextChange | null = null;
  for (const segment of segments) {
    if (segment.type === "equal") {
      if (pending !== null) {
        changes.push(pending);
        pending = null;
      }
      cursor += segment.text.length;
      continue;
    }
    pending ??= { start: cursor, end: cursor, text: "" };
    if (segment.type === "del") {
      cursor += segment.text.length;
      pending.end = cursor;
    } else {
      pending.text += segment.text;
    }
  }
  if (pending !== null) {
    changes.push(pending);
  }
  return changes;
};

const cutsInto = ({ offset, length }: AtomicTextSpan, position: number): boolean =>
  position > offset && position < offset + length;

/**
 * `changes` widened so none cuts into an atomic span, then merged where the
 * widening made two overlap. A widened change re-emits the source text it
 * absorbed, so applying the result still yields exactly the replacement.
 *
 * `spans` are in the same coordinates as the changes (the source span's).
 */
export const widenChangesToAtomicSpans = (
  source: string,
  changes: readonly TextChange[],
  spans: readonly AtomicTextSpan[],
): TextChange[] => {
  if (spans.length === 0 || changes.length === 0) {
    return [...changes];
  }
  const windows = changes.map(({ start, end }) => {
    let from = start;
    let to = end;
    for (const span of spans) {
      if (cutsInto(span, from)) {
        from = span.offset;
      }
      if (cutsInto(span, to)) {
        to = span.offset + span.length;
      }
    }
    return { from, to };
  });

  const merged: TextChange[] = [];
  let index = 0;
  while (index < changes.length) {
    // SAFETY: `index` is bounded by `changes.length`, and `windows` has one
    // entry per change.
    const { from } = windows[index]!;
    let { to } = windows[index]!;
    const group: TextChange[] = [changes[index]!];
    index++;
    while (index < changes.length && windows[index]!.from < to) {
      to = Math.max(to, windows[index]!.to);
      group.push(changes[index]!);
      index++;
    }
    let text = "";
    let cursor = from;
    for (const change of group) {
      text += source.slice(cursor, change.start) + change.text;
      cursor = change.end;
    }
    text += source.slice(cursor, to);
    merged.push({ start: from, end: to, text });
  }
  return merged;
};

/** `source` with `changes` applied; the invariant the planner owes. */
export const applyTextChanges = (source: string, changes: readonly TextChange[]): string => {
  let result = "";
  let cursor = 0;
  for (const change of changes) {
    result += source.slice(cursor, change.start) + change.text;
    cursor = change.end;
  }
  return result + source.slice(cursor);
};
