/**
 * Footnote and endnote references in the AI text view.
 *
 * The editor holds a note reference as a text node carrying a `footnoteRef`
 * mark whose characters are the note's package `w:id` (`10`, `30`). Those
 * characters are not document text: the page shows the note's reading-order number
 * there, and the save writes a `w:footnoteReference` whatever the node holds.
 * Read as text, an id looks like prose (`Term A10`), a search for `10` finds
 * it, and a replacement of it silently deletes the reference.
 *
 * So every AI-facing reader shows a reference as its marker — `[^1]` for the
 * first footnote, `[^e1]` for the first endnote, the numbering the Markdown
 * export uses (`utils/noteReferenceLabels`) — and the clean-text builder
 * records it as an atomic structural span that no text range may cut into and
 * no text change may alter.
 */

import type { Mark, Node as PMNode } from "prosemirror-model";

import {
  createNoteReferenceNumbering,
  noteReferenceMarker,
  type NoteReferenceKind,
} from "../utils/noteReferenceLabels";
import type { CleanBlockText } from "./clean-text";
import type { TextChange } from "./minimal-replacement";

const NOTE_REFERENCE_MARK = "footnoteRef";
const DELETION_MARK = "deletion";
const HIDDEN_MARK = "hidden";

/** Which note a reference names: its kind and its package id. */
export type NoteReferenceIdentity = {
  noteType: NoteReferenceKind;
  noteId: string;
};

/** The reference a text node carries, or `null` for ordinary text. */
export const noteReferenceOf = (node: PMNode): NoteReferenceIdentity | null => {
  if (!node.isText) {
    return null;
  }
  const mark = node.marks.find((candidate: Mark) => candidate.type.name === NOTE_REFERENCE_MARK);
  if (!mark) {
    return null;
  }
  return {
    noteType: mark.attrs["noteType"] === "endnote" ? "endnote" : "footnote",
    noteId: String(mark.attrs["id"]),
  };
};

/**
 * The marker each reference of a story shows, keyed by note: numbered once
 * for a whole story so every block reads the same number the Markdown export
 * gives it, and so an edit elsewhere in one batch cannot renumber a block the
 * batch resolved earlier.
 */
export type NoteReferenceLabels = {
  labelOf: (reference: NoteReferenceIdentity) => string | undefined;
};

const labelKey = ({ noteType, noteId }: NoteReferenceIdentity): string => `${noteType}:${noteId}`;

const isOmittedFromCleanView = (node: PMNode): boolean =>
  node.marks.some((mark) => mark.type.name === DELETION_MARK || mark.type.name === HIDDEN_MARK);

/** Labels that grow as the blocks of a story are numbered, in reading order. */
export type NoteReferenceLabeler = NoteReferenceLabels & {
  /**
   * Number the references of one block. A reference the clean view omits
   * (tracked-deleted, hidden) takes no number, exactly as its characters take
   * no place in the clean text.
   */
  numberBlock: (block: PMNode) => void;
};

/**
 * The one numbering every AI reader of a story shares. The snapshot numbers
 * block by block inside its own walk; `collectNoteReferenceLabels` (with the
 * snapshot, so both skip what the snapshot skips) numbers a whole story.
 */
export const createNoteReferenceLabeler = (): NoteReferenceLabeler => {
  const labels = new Map<string, string>();
  const numbering = createNoteReferenceNumbering();
  return {
    labelOf: (reference) => labels.get(labelKey(reference)),
    numberBlock: (block) => {
      block.descendants((node) => {
        const reference = noteReferenceOf(node);
        if (reference === null || isOmittedFromCleanView(node)) {
          return true;
        }
        const { displayNumber, first } = numbering.next(reference.noteType, reference.noteId);
        if (first) {
          labels.set(labelKey(reference), noteReferenceMarker(reference.noteType, displayNumber));
        }
        return true;
      });
    },
  };
};

/**
 * The text a reader shows for an inline text node: a note reference's marker
 * instead of its package id. A reference the clean view omits (a tracked
 * deletion) has no number there, and reads as an unnumbered marker.
 */
export const readerTextOf = (node: PMNode, labels: NoteReferenceLabels): string => {
  const reference = noteReferenceOf(node);
  if (reference === null) {
    return node.text ?? "";
  }
  return (
    labels.labelOf(reference) ??
    (reference.noteType === "endnote" ? UNNUMBERED_ENDNOTE_MARKER : UNNUMBERED_FOOTNOTE_MARKER)
  );
};

const UNNUMBERED_FOOTNOTE_MARKER = "[^?]";
const UNNUMBERED_ENDNOTE_MARKER = "[^e?]";

/** A note reference's marker inside a span of clean text, relative to the span. */
export type NoteReferenceSpan = { offset: number; length: number };

type StructuralBoundaries = Pick<CleanBlockText, "structuralBoundaries">;

/** The markers lying wholly inside clean text `[start, end)`, relative to `start`. */
export const noteReferenceSpansWithin = (
  { structuralBoundaries }: StructuralBoundaries,
  start: number,
  end: number,
): NoteReferenceSpan[] =>
  structuralBoundaries.flatMap((boundary) =>
    boundary.type === "noteReference" &&
    boundary.offset >= start &&
    boundary.offset + boundary.length <= end
      ? [{ offset: boundary.offset - start, length: boundary.length }]
      : [],
  );

/** Whether clean text `[start, end)` starts or ends inside a reference's marker. */
export const cutsIntoNoteReference = (
  { structuralBoundaries }: StructuralBoundaries,
  start: number,
  end: number,
): boolean =>
  structuralBoundaries.some(
    (boundary) =>
      boundary.type === "noteReference" &&
      [start, end].some(
        (offset) => offset > boundary.offset && offset < boundary.offset + boundary.length,
      ),
  );

let comparisonEditsNoteReferences = false;

/**
 * @internal Run `apply` with its text operations free to remove note
 * references and to write marker text in place of new ones.
 *
 * Only a document comparison does this. Its operations come from a revised
 * document that adds and removes references, so refusing them would drop the
 * edit; each marker it writes is turned back into the reference the revised
 * document holds once the operations land (`compare/inline-atoms.ts`), and a
 * removed reference is a tracked deletion of the reference itself. Every other
 * caller keeps the protection. The scope is synchronous, like the applier.
 */
export const withComparisonNoteReferenceEdits = <T>(apply: () => T): T => {
  const previous = comparisonEditsNoteReferences;
  comparisonEditsNoteReferences = true;
  try {
    return apply();
  } finally {
    comparisonEditsNoteReferences = previous;
  }
};

/** @internal Whether the running apply belongs to a comparison (see above). */
export const noteReferenceEditsAllowed = (): boolean => comparisonEditsNoteReferences;

const NOTE_MARKER_PATTERN = /\[\^e?[1-9]\d*\]/gu;
const markerCount = (text: string): number => text.match(NOTE_MARKER_PATTERN)?.length ?? 0;

/** The pieces of a replacement between the markers it has to keep. */
type NoteReferenceSegments = {
  /** The source text between markers, with its offset in the source. */
  source: { offset: number; text: string }[];
  /** What each source piece becomes. */
  replacement: string[];
};

/**
 * Pair a text replacement with the reference markers its source holds.
 *
 * A reference is not text, so a replacement may only rewrite the prose around
 * it: every marker of `source` has to reappear in `replacement`, in order, and
 * the prose between two markers becomes the prose between the same two. A
 * replacement that drops, reorders or rewrites a marker — or writes a new
 * marker-shaped string, which would read as a reference and be none — answers
 * `null`: the edit is refused rather than deleting a reference, or leaving
 * text behind that claims to be one.
 */
export const segmentsAroundNoteReferences = (
  source: string,
  replacement: string,
  spans: readonly NoteReferenceSpan[],
): NoteReferenceSegments | null => {
  const segments: NoteReferenceSegments = { source: [], replacement: [] };
  let sourceCursor = 0;
  let replacementCursor = 0;
  for (const { offset, length } of spans) {
    const marker = source.slice(offset, offset + length);
    const at = replacement.indexOf(marker, replacementCursor);
    if (at === -1) {
      return null;
    }
    segments.source.push({ offset: sourceCursor, text: source.slice(sourceCursor, offset) });
    segments.replacement.push(replacement.slice(replacementCursor, at));
    sourceCursor = offset + length;
    replacementCursor = at + marker.length;
  }
  segments.source.push({ offset: sourceCursor, text: source.slice(sourceCursor) });
  segments.replacement.push(replacement.slice(replacementCursor));
  for (const [index, piece] of segments.replacement.entries()) {
    if (markerCount(piece) > markerCount(segments.source[index]?.text ?? "")) {
      return null;
    }
  }
  return segments;
};

/**
 * Plan a replacement's changes piece by piece between the markers it keeps, so
 * no change ever touches a reference. `null` when the replacement does not
 * keep them (see {@link segmentsAroundNoteReferences}).
 */
export const planChangesAroundNoteReferences = (
  source: string,
  replacement: string,
  spans: readonly NoteReferenceSpan[],
  planChanges: (source: string, replacement: string) => readonly TextChange[],
): TextChange[] | null => {
  const segments = segmentsAroundNoteReferences(source, replacement, spans);
  if (segments === null) {
    return null;
  }
  return segments.source.flatMap(({ offset, text }, index) =>
    planChanges(text, segments.replacement[index] ?? "").map((change) => ({
      start: change.start + offset,
      end: change.end + offset,
      text: change.text,
    })),
  );
};
