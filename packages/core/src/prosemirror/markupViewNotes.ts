/**
 * The notes a markup view shows: each note story with its revisions resolved
 * the way the view resolves the body's. Original shows every note as it was
 * before the revisions (a note whose reference's deletion takes its text
 * shows that text), No Markup and Simple Markup as it reads once they are
 * accepted. All Markup shows the notes as authored.
 *
 * Resolved notes are cached per note and view mode, so an unchanged note is
 * not resolved again on every layout pass.
 */

import type { EditorState } from "prosemirror-state";

import {
  findStoryBody,
  resolvedDocumentView,
  type RevisionDecision,
} from "@stll/docx-core/ops";

import type { DisplayMode } from "../managers/EditorModeManager";
import type { Document, Endnote, Footnote } from "../types/document";
import { markupViewResolutionMode } from "./markupViewProjection";

type NoteReviewContext = Pick<
  Document["package"],
  "styles" | "theme" | "numbering"
>;
type ResolvedNote = NoteReviewContext & { note: Footnote | Endnote };
const resolvedNotes = new WeakMap<
  Footnote | Endnote,
  Map<RevisionDecision, ResolvedNote>
>();

const resolvedNote = <TNote extends Footnote | Endnote>(
  note: TNote,
  decision: RevisionDecision,
  { styles, theme, numbering }: NoteReviewContext,
): TNote => {
  const byDecision =
    resolvedNotes.get(note) ?? new Map<RevisionDecision, ResolvedNote>();
  const cached = byDecision.get(decision);
  if (
    cached !== undefined &&
    cached.styles === styles &&
    cached.theme === theme &&
    cached.numbering === numbering
  )
    return cached.note as TNote;

  // View resolution must retain the note even if resolving the body would
  // discard its reference. The isolated package addresses the real note story.
  const source = {
    package: {
      document: { content: [] },
      ...(note.type === "footnote"
        ? { footnotes: [note] }
        : { endnotes: [note] }),
      ...(styles !== undefined && { styles }),
      ...(theme !== undefined && { theme }),
      ...(numbering !== undefined && { numbering }),
    },
  } satisfies Document;
  const view = resolvedDocumentView(source, decision);
  const body = findStoryBody(view.document, { kind: note.type, id: note.id });
  const resolved =
    body === undefined || body.content === note.content
      ? note
      : { ...note, content: body.content };
  byDecision.set(decision, {
    ...(styles !== undefined && { styles }),
    ...(theme !== undefined && { theme }),
    ...(numbering !== undefined && { numbering }),
    note: resolved,
  });
  resolvedNotes.set(note, byDecision);
  return resolved;
};

/** The footnotes or endnotes as `view` shows them. */
export const markupViewNotes = <TNote extends Footnote | Endnote>(
  notes: TNote[] | undefined,
  view: DisplayMode,
  _state: EditorState,
  document: Document | null,
): TNote[] | undefined => {
  const decision = markupViewResolutionMode(view);
  if (!notes || !decision || !document) return notes;
  let changed = false;
  const shown = notes.map((note) => {
    const resolved = resolvedNote(note, decision, document.package);
    changed ||= resolved !== note;
    return resolved;
  });
  return changed ? shown : notes;
};
