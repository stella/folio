/**
 * A note goes with its reference. Deleting a note's reference while tracking
 * changes deletes the note's text too, so rejecting that deletion must give
 * the note its text back, whichever surface resolves it: the body editor
 * resolves the reference, and this resolves the note's story to match.
 */

import type { Mark, Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";
import { Transform } from "prosemirror-transform";

import { withoutLostCommentThreads } from "../ai-edits/comment-lifecycle";
import { visitCommentMarkers } from "../docx/commentAnchorIndex";
import { isSeparatorEndnote, isSeparatorFootnote } from "../docx/footnoteParser";
import { pluginsForHeadlessRevisionResolution } from "../internal/headlessRevisionResolutionGuard";
import { expectFootnoteRefMarkAttrs, expectParagraphAttrs } from "./attrs";
import { rejectAIEditRevision } from "./commands/comments";
import { proseDocToBlocks } from "./conversion/fromProseDoc";
import { footnoteToProseDoc } from "./conversion/toProseDoc";
import { createDocumentNumberingPlugin } from "./plugins/documentNumbering";
import { createDocumentStylesPlugin } from "./plugins/documentStyles";
import { mintRevisionId } from "./plugins/revisionIds";
import { schema, singletonManager } from "./schema";
import type { BlockContent, Document, Endnote, Footnote } from "../types/document";

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

/** The deletion each pending-deleted note reference carries. */
const referenceDeletions = (doc: PMNode): Map<NoteKey, Mark> => {
  const deletions = new Map<NoteKey, Mark>();
  doc.descendants((node) => {
    if (!node.isText) return true;
    const reference = node.marks.find((mark) => mark.type.name === "footnoteRef");
    const deletion = node.marks.find((mark) => mark.type.name === "deletion");
    if (reference && deletion) {
      const { id, noteType } = expectFootnoteRefMarkAttrs(reference);
      deletions.set(noteKey(noteType === "endnote" ? "endnote" : "footnote", Number(id)), deletion);
    }
    return false;
  });
  return deletions;
};

/**
 * A note whose reference was deleted: its text and its paragraph marks are
 * deleted too, by the reference's author at the reference's date.
 */
const deletedNote = <TNote extends Footnote | Endnote>(
  note: TNote,
  document: Document,
  reference: Mark,
): TNote => {
  const { styles, theme } = document.package;
  const doc = footnoteToProseDoc(note.content, {
    ...(styles !== undefined && { styles }),
    ...(theme !== undefined && { theme }),
  });
  const deletion = reference.type.create({ ...reference.attrs, revisionId: mintRevisionId() });
  const markInfo = {
    id: mintRevisionId(),
    author: String(reference.attrs["author"] ?? ""),
    date: String(reference.attrs["date"] ?? ""),
  };
  const transform = new Transform(doc);
  doc.descendants((node, position) => {
    if (node.type.name === "paragraph" && expectParagraphAttrs(node).pPrMark == null) {
      transform.setNodeAttribute(position, "pPrMark", { kind: "del", info: markInfo });
    }
    if (node.isInline && !deletion.type.isInSet(node.marks)) {
      transform.addMark(position, position + node.nodeSize, deletion);
    }
    return !node.isInline;
  });
  if (!transform.docChanged) return note;
  return { ...note, content: proseDocToBlocks(transform.doc, note.content, styles) };
};

type NoteEdit = (note: Footnote | Endnote) => Footnote | Endnote;

const editNotes = (document: Document, edits: ReadonlyMap<NoteKey, NoteEdit>): Document => {
  if (edits.size === 0) return document;
  const edit = <TNote extends Footnote | Endnote>(
    kind: NoteKind,
    notes: TNote[] | undefined,
  ): TNote[] | undefined => {
    if (!notes?.some((note) => edits.has(noteKey(kind, note.id)))) return notes;
    return notes.map((note) => (edits.get(noteKey(kind, note.id))?.(note) as TNote) ?? note);
  };
  const footnotes = edit("footnote", document.package.footnotes);
  const endnotes = edit("endnote", document.package.endnotes);
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

/**
 * Keeps an editor's notes following their references across body edits. The
 * editor reports the body as it was before the edits it has not yet written
 * back, and reconciles the document it writes back:
 *
 * - a reference whose deletion is rejected gives its note back its text;
 * - a reference deleted again, as undoing that reject does, takes the note's
 *   text with it: the note returns to what it was, or, when the note changed
 *   since, its text is deleted along with the reference.
 */
export type NoteReferenceFollower = {
  /** The body before the edits the next {@link reconcile} writes back. */
  noteBase: (before: PMNode) => void;
  /** The document to write back, its notes following the body's references. */
  reconcile: (document: Document, body: PMNode) => Document;
  /** Forget the pending base and what earlier restores replaced (a new document). */
  reset: () => void;
};

export const createNoteReferenceFollower = (): NoteReferenceFollower => {
  let base: PMNode | null = null;
  // What each restore replaced, while the restored note stays as restored.
  const replaced = new Map<
    NoteKey,
    { restored: Footnote | Endnote; previous: Footnote | Endnote }
  >();
  return {
    noteBase: (before) => {
      base ??= before;
    },
    reconcile: (document, body) => {
      const before = base;
      base = null;
      if (!before) return document;
      const was = noteReferenceStates(before);
      const now = noteReferenceStates(body);
      const deletions = referenceDeletions(body);
      const edits = new Map<NoteKey, NoteEdit>();
      for (const [key, state] of now) {
        const previous = was.get(key);
        if (state === "live" && previous === "deleted") {
          edits.set(key, (note) => {
            const restored = restoredNote(note, document);
            if (restored !== note) replaced.set(key, { restored, previous: note });
            return restored;
          });
        }
        const reference = deletions.get(key);
        if (state === "deleted" && previous === "live" && reference) {
          edits.set(key, (note) => {
            const earlier = replaced.get(key);
            replaced.delete(key);
            return earlier?.restored === note
              ? earlier.previous
              : deletedNote(note, document, reference);
          });
        }
      }
      return editNotes(document, edits);
    },
    reset: () => {
      base = null;
      replaced.clear();
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

/** Ids of the comments anchored anywhere in these stories. */
const anchoredComments = (stories: readonly (readonly BlockContent[])[]): Set<number> => {
  const ids = new Set<number>();
  for (const blocks of stories) visitCommentMarkers(blocks, ({ item }) => ids.add(item.id));
  return ids;
};

/**
 * The document without the notes nothing refers to any more: a note goes
 * with its reference, so accepting the reference's deletion leaves a note no
 * reader can reach, and saving drops it, with the comments anchored only in
 * it. Separators stay.
 */
export const withoutUnreferencedNotes = (document: Document): Document => {
  const referenced = referencedNotes(document);
  const { footnotes, endnotes, headers, footers } = document.package;
  const keeps = (kind: NoteKind) => (note: Footnote | Endnote) =>
    (kind === "footnote" ? isSeparatorFootnote(note) : isSeparatorEndnote(note)) ||
    referenced.has(noteKey(kind, note.id));
  const keptFootnotes = footnotes?.filter(keeps("footnote"));
  const keptEndnotes = endnotes?.filter(keeps("endnote"));
  if (keptFootnotes?.length === footnotes?.length && keptEndnotes?.length === endnotes?.length) {
    return document;
  }
  const removed = [
    ...(footnotes ?? []).filter((note) => !keeps("footnote")(note)),
    ...(endnotes ?? []).filter((note) => !keeps("endnote")(note)),
  ];
  const comments = document.package.document.comments;
  const keptComments = comments
    ? withoutLostCommentThreads(comments, {
        anchoredBefore: anchoredComments(removed.map((note) => note.content)),
        anchoredNow: anchoredComments([
          document.package.document.content,
          ...[...(headers?.values() ?? []), ...(footers?.values() ?? [])].map(
            (part) => part.content,
          ),
          ...[...(keptFootnotes ?? []), ...(keptEndnotes ?? [])].map((note) => note.content),
        ]),
      })
    : undefined;
  return {
    ...document,
    package: {
      ...document.package,
      ...(keptFootnotes !== undefined && { footnotes: keptFootnotes }),
      ...(keptEndnotes !== undefined && { endnotes: keptEndnotes }),
      ...(keptComments !== undefined &&
        keptComments.length !== comments?.length && {
          document: { ...document.package.document, comments: keptComments },
        }),
    },
  };
};
