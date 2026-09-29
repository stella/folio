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

import { resolveWholeStory } from "../internal/wholeStoryRevisionResolution";
import type { RevisionResolutionMode } from "../internal/revisionResolutionInline";
import type { DisplayMode } from "../managers/EditorModeManager";
import { proseDocToBlocks } from "./conversion/fromProseDoc";
import { footnoteToProseDoc } from "./conversion/toProseDoc";
import { markupViewResolutionMode } from "./markupViewProjection";
import { getDocumentNumbering } from "./plugins/documentNumbering";
import { getDocumentStyleResolver } from "./plugins/documentStyles";
import type { Document, Endnote, Footnote } from "../types/document";

type ResolvedNote = { styleResolver: unknown; note: Footnote | Endnote };
const resolvedNotes = new WeakMap<object, Map<RevisionResolutionMode, ResolvedNote>>();

const resolvedNote = <TNote extends Footnote | Endnote>(
  note: TNote,
  mode: RevisionResolutionMode,
  state: EditorState,
  document: Document,
): TNote => {
  const styleResolver = getDocumentStyleResolver(state);
  const cached = resolvedNotes.get(note)?.get(mode);
  if (cached?.styleResolver === styleResolver) return cached.note as TNote;
  const { styles, theme } = document.package;
  const doc = footnoteToProseDoc(note.content, {
    ...(styles !== undefined && { styles }),
    ...(theme !== undefined && { theme }),
  });
  const result = resolveWholeStory({
    doc,
    mode,
    styleResolver,
    numbering: getDocumentNumbering(state),
  });
  const resolved =
    result.failed || result.resolved.eq(doc)
      ? note
      : { ...note, content: proseDocToBlocks(result.resolved, note.content, styles) };
  const byMode = resolvedNotes.get(note) ?? new Map<RevisionResolutionMode, ResolvedNote>();
  byMode.set(mode, { styleResolver, note: resolved });
  resolvedNotes.set(note, byMode);
  return resolved;
};

/** The footnotes or endnotes as `view` shows them. */
export const markupViewNotes = <TNote extends Footnote | Endnote>(
  notes: TNote[] | undefined,
  view: DisplayMode,
  state: EditorState,
  document: Document | null,
): TNote[] | undefined => {
  const mode = markupViewResolutionMode(view);
  if (!notes || !mode || !document) return notes;
  let changed = false;
  const shown = notes.map((note) => {
    const resolved = resolvedNote(note, mode, state, document);
    changed ||= resolved !== note;
    return resolved;
  });
  return changed ? shown : notes;
};
