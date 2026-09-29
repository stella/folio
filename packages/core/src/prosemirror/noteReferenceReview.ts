/**
 * A note goes with its reference. Deleting a note's reference while tracking
 * changes deletes the note's text too, so rejecting that deletion must give
 * the note its text back, whichever surface resolves it: the body editor
 * resolves the reference, and this resolves the note's story to match.
 */

import type { Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";

import { isSeparatorEndnote, isSeparatorFootnote } from "../docx/footnoteParser";
import { pluginsForHeadlessRevisionResolution } from "../internal/headlessRevisionResolutionGuard";
import { expectFootnoteRefMarkAttrs, expectParagraphAttrs } from "./attrs";
import { rejectAIEditRevision } from "./commands/comments";
import { proseDocToBlocks } from "./conversion/fromProseDoc";
import { footnoteToProseDoc } from "./conversion/toProseDoc";
import { createDocumentNumberingPlugin } from "./plugins/documentNumbering";
import { createDocumentStylesPlugin } from "./plugins/documentStyles";
import { schema, singletonManager } from "./schema";
import type { Document, Endnote, Footnote } from "../types/document";

type NoteKind = "footnote" | "endnote";

/** `footnote:<id>` or `endnote:<id>`. */
export type NoteKey = `${NoteKind}:${number}`;

const noteKey = (kind: NoteKind, id: number): NoteKey => `${kind}:${id}`;

/**
 * The notes a story's text refers to, and whether each reference is pending
 * deletion: a reference is deleted when every character of it is.
 */
export const noteReferenceStates = (doc: PMNode): Map<NoteKey, "live" | "deleted"> => {
  const states = new Map<NoteKey, "live" | "deleted">();
  doc.descendants((node) => {
    if (!node.isText) return true;
    const reference = node.marks.find((mark) => mark.type.name === "footnoteRef");
    if (!reference) return false;
    const { id, noteType } = expectFootnoteRefMarkAttrs(reference);
    const key = noteKey(noteType === "endnote" ? "endnote" : "footnote", Number(id));
    const deleted = node.marks.some((mark) => mark.type.name === "deletion");
    if (!deleted || states.get(key) === "live") {
      states.set(key, "live");
    } else {
      states.set(key, "deleted");
    }
    return false;
  });
  return states;
};

/** The notes whose reference a change took from pending deletion back to live. */
export const noteReferencesRestored = (before: PMNode, after: PMNode): NoteKey[] => {
  const was = noteReferenceStates(before);
  const now = noteReferenceStates(after);
  return [...now].flatMap(([key, state]) =>
    state === "live" && was.get(key) === "deleted" ? [key] : [],
  );
};

/** Every revision that deletes part of a note: its text and its paragraph marks. */
export const noteDeletionRevisions = (doc: PMNode): number[] => {
  const ids = new Set<number>();
  doc.descendants((node) => {
    if (node.type.name === "paragraph") {
      const mark = expectParagraphAttrs(node).pPrMark;
      if (mark?.kind === "del") ids.add(mark.info.id);
    }
    for (const mark of node.marks) {
      if (mark.type.name === "deletion") {
        const id: unknown = mark.attrs["revisionId"];
        if (typeof id === "number") ids.add(id);
      }
    }
    return true;
  });
  return [...ids];
};

const restoredNote = <TNote extends Footnote | Endnote>(note: TNote, document: Document): TNote => {
  const { styles, theme, numbering } = document.package;
  const doc = footnoteToProseDoc(note.content, {
    ...(styles !== undefined && { styles }),
    ...(theme !== undefined && { theme }),
  });
  const revisions = noteDeletionRevisions(doc);
  if (revisions.length === 0) return note;
  let state = EditorState.create({
    schema,
    doc,
    plugins: [
      ...pluginsForHeadlessRevisionResolution(singletonManager.getPlugins()),
      createDocumentStylesPlugin(styles),
      createDocumentNumberingPlugin(numbering),
    ],
  });
  rejectAIEditRevision(revisions)(state, (transaction) => {
    state = state.apply(transaction);
  });
  if (state.doc.eq(doc)) return note;
  return { ...note, content: proseDocToBlocks(state.doc, note.content, styles) };
};

/**
 * The document with the deletions in the named notes rejected: the notes
 * whose reference's deletion was rejected get their text back.
 */
export const restoreNotes = (document: Document, keys: Iterable<NoteKey>): Document => {
  const wanted = new Set(keys);
  if (wanted.size === 0) return document;
  const restore = <TNote extends Footnote | Endnote>(
    kind: NoteKind,
    notes: TNote[] | undefined,
  ): TNote[] | undefined => {
    if (!notes?.some((note) => wanted.has(noteKey(kind, note.id)))) return notes;
    return notes.map((note) =>
      wanted.has(noteKey(kind, note.id)) ? restoredNote(note, document) : note,
    );
  };
  const footnotes = restore("footnote", document.package.footnotes);
  const endnotes = restore("endnote", document.package.endnotes);
  if (footnotes === document.package.footnotes && endnotes === document.package.endnotes) {
    return document;
  }
  return {
    ...document,
    package: {
      ...document.package,
      ...(footnotes !== undefined && { footnotes }),
      ...(endnotes !== undefined && { endnotes }),
    },
  };
};

/** The notes the main story refers to, live or pending deletion. */
const referencedNotes = (document: Document): Set<NoteKey> => {
  const keys = new Set<NoteKey>();
  const visit = (node: unknown): void => {
    if (Array.isArray(node)) {
      for (const item of node) visit(item);
      return;
    }
    if (node === null || typeof node !== "object") return;
    const { type, id } = node as { type?: unknown; id?: unknown };
    if ((type === "footnoteRef" || type === "endnoteRef") && typeof id === "number") {
      keys.add(noteKey(type === "footnoteRef" ? "footnote" : "endnote", id));
    }
    for (const value of Object.values(node)) visit(value);
  };
  visit(document.package.document.content);
  return keys;
};

/**
 * The document without the notes nothing refers to any more: a note goes
 * with its reference, so accepting the reference's deletion leaves a note no
 * reader can reach, and saving drops it. Separators stay.
 */
export const withoutUnreferencedNotes = (document: Document): Document => {
  const referenced = referencedNotes(document);
  const { footnotes, endnotes } = document.package;
  const keptFootnotes = footnotes?.filter(
    (note) => isSeparatorFootnote(note) || referenced.has(noteKey("footnote", note.id)),
  );
  const keptEndnotes = endnotes?.filter(
    (note) => isSeparatorEndnote(note) || referenced.has(noteKey("endnote", note.id)),
  );
  if (keptFootnotes?.length === footnotes?.length && keptEndnotes?.length === endnotes?.length) {
    return document;
  }
  return {
    ...document,
    package: {
      ...document.package,
      ...(keptFootnotes !== undefined && { footnotes: keptFootnotes }),
      ...(keptEndnotes !== undefined && { endnotes: keptEndnotes }),
    },
  };
};
