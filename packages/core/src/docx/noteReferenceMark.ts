/**
 * The run that shows a note's own number at the start of its text
 * (`w:footnoteRef` / `w:endnoteRef`).
 *
 * A parsed note carries it as a typed run child, so its revisions — a
 * deletion of the whole note, say — wrap it like any other run. A note made
 * in the editor gets the same run when it is created, rather than having the
 * save add one outside everything the note's content says about itself.
 */

import type { BlockContent, Endnote, Footnote, Run } from "../types/document";
import { cloneParagraphWithPropertySource } from "./paragraphPropertySource";
import { visitDocxParagraphs, visitParagraphRuns } from "./paragraphTraversal";

type NoteKind = "footnote" | "endnote";

/** The run carrying a note's in-note reference mark, in the note reference style. */
export const noteReferenceMarkRun = (kind: NoteKind): Run => ({
  type: "run",
  formatting: { styleId: kind === "footnote" ? "FootnoteReference" : "EndnoteReference" },
  content: [{ type: "noteMarker", kind }],
});

/** Whether a note already owns its typed marker, including inside tracked content. */
export const hasNoteReferenceMark = (kind: NoteKind, content: readonly BlockContent[]): boolean => {
  let found = false;
  visitDocxParagraphs({ documentBody: { content: [...content] } }, (paragraph) => {
    visitParagraphRuns(paragraph, (run) => {
      if (run.content.some((child) => child.type === "noteMarker" && child.kind === kind)) {
        found = true;
      }
    });
  });
  return found;
};

/**
 * A note's content with its reference mark run at the start of its first
 * paragraph, added when the note has none.
 */
export const withNoteReferenceMark = (
  kind: NoteKind,
  content: readonly BlockContent[],
): BlockContent[] => {
  if (hasNoteReferenceMark(kind, content)) {
    return [...content];
  }
  const index = content.findIndex((block) => block.type === "paragraph");
  const mark = noteReferenceMarkRun(kind);
  if (index === -1) {
    return [{ type: "paragraph", content: [mark] }, ...content];
  }
  return content.map((block, position) =>
    position === index && block.type === "paragraph"
      ? cloneParagraphWithPropertySource(block, { content: [mark, ...block.content] })
      : block,
  );
};

/** A new note, as an editor adds one: its content led by its own reference mark. */
export function createNote(
  kind: "footnote",
  id: number,
  content: readonly BlockContent[],
): Footnote;
export function createNote(kind: "endnote", id: number, content: readonly BlockContent[]): Endnote;
export function createNote(
  kind: NoteKind,
  id: number,
  content: readonly BlockContent[],
): Footnote | Endnote {
  return kind === "footnote"
    ? { type: "footnote", id, content: withNoteReferenceMark(kind, content) }
    : { type: "endnote", id, content: withNoteReferenceMark(kind, content) };
}
