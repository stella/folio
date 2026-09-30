/**
 * The run that shows a note's own number at the start of its text
 * (`w:footnoteRef` / `w:endnoteRef`).
 *
 * A parsed note carries it as a preserved run child, so its revisions — a
 * deletion of the whole note, say — wrap it like any other run. A note made
 * in the editor gets the same run when it is created, rather than having the
 * save add one outside everything the note's content says about itself.
 */

import type { BlockContent, Endnote, Footnote, Paragraph, Run } from "../types/document";
import { cloneParagraphWithPropertySource } from "./paragraphPropertySource";

type NoteKind = "footnote" | "endnote";

const MARK_ELEMENT = /<(?:[\w.-]+:)?(?:footnote|endnote)Ref[\s/>]/u;

/** The run carrying a note's in-note reference mark, in the note reference style. */
export const noteReferenceMarkRun = (kind: NoteKind): Run => ({
  type: "run",
  formatting: { styleId: kind === "footnote" ? "FootnoteReference" : "EndnoteReference" },
  content: [{ type: "preservedXml", xml: `<w:${kind}Ref/>`, text: "" }],
});

const holdsReferenceMark = (value: unknown): boolean => {
  if (Array.isArray(value)) {
    return value.some(holdsReferenceMark);
  }
  if (typeof value !== "object" || value === null) {
    return false;
  }
  const record = value as Record<string, unknown>;
  if (record["type"] === "preservedXml" && typeof record["xml"] === "string") {
    return MARK_ELEMENT.test(record["xml"]);
  }
  return Object.values(record).some(holdsReferenceMark);
};

/**
 * A note's content with its reference mark run at the start of its first
 * paragraph, added when the note has none.
 */
export const withNoteReferenceMark = (
  kind: NoteKind,
  content: readonly BlockContent[],
): BlockContent[] => {
  if (holdsReferenceMark(content)) {
    return [...content];
  }
  const index = content.findIndex((block) => block.type === "paragraph");
  const mark = noteReferenceMarkRun(kind);
  if (index === -1) {
    return [{ type: "paragraph", content: [mark] }, ...content];
  }
  const paragraph = content[index] as Paragraph;
  return content.map((block, position) =>
    position === index
      ? cloneParagraphWithPropertySource(paragraph, { content: [mark, ...paragraph.content] })
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
