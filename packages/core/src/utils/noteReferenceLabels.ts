/**
 * How a reader shows a footnote or endnote reference.
 *
 * A reference's `w:id` is a package key, not something a reader sees: the
 * rendered document numbers notes by the order they are first referenced
 * (1, 2, 3 …), whatever ids the package gave them. Every reader that puts a reference in a line of
 * text — the Markdown export, `getContent()`, the AI snapshot and the agent
 * tools built on it — numbers through {@link createNoteReferenceNumbering}, so
 * the same reference reads as the same marker everywhere.
 *
 * Footnotes and endnotes are numbered in separate sequences; an
 * endnote's token carries an `e` so the two sequences never collide
 * (`[^1]` is footnote 1, `[^e1]` is endnote 1).
 */

export type NoteReferenceKind = "footnote" | "endnote";

/** The bare token: `1` for footnote 1, `e1` for endnote 1. */
export const noteReferenceToken = (kind: NoteReferenceKind, displayNumber: number): string =>
  kind === "endnote" ? `e${displayNumber}` : String(displayNumber);

/** The marker a text reader shows in the line: `[^1]`, `[^e1]`. */
export const noteReferenceMarker = (kind: NoteReferenceKind, displayNumber: number): string =>
  `[^${noteReferenceToken(kind, displayNumber)}]`;

export type NoteReferenceNumber = {
  displayNumber: number;
  /** Whether this is the note's first reference in reading order. */
  first: boolean;
};

export type NoteReferenceNumbering = {
  /** Number one reference, in reading order. */
  next: (kind: NoteReferenceKind, noteId: string | number) => NoteReferenceNumber;
};

/**
 * A reading-order numbering: each note takes the next number of its kind the
 * first time it is referenced, and a repeated reference reuses it.
 */
export const createNoteReferenceNumbering = (): NoteReferenceNumbering => {
  const numbers = new Map<string, number>();
  const counts: Record<NoteReferenceKind, number> = { footnote: 0, endnote: 0 };
  return {
    next: (kind, noteId) => {
      const key = `${kind}:${String(noteId)}`;
      const existing = numbers.get(key);
      if (existing !== undefined) {
        return { displayNumber: existing, first: false };
      }
      counts[kind] += 1;
      numbers.set(key, counts[kind]);
      return { displayNumber: counts[kind], first: true };
    },
  };
};
