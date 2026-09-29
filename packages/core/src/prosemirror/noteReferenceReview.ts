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

/**
 * The revisions that deleted a note along with its reference: its text and
 * its paragraph marks, deleted by the reference deletion's author at its
 * date. A deletion in the note made some other time is a change of its own,
 * which resolving the reference leaves alone.
 */
export const noteDeletionRevisions = (doc: PMNode, reference: Mark): number[] => {
  const author = String(reference.attrs["author"] ?? "");
  const date = String(reference.attrs["date"] ?? "");
  const ids = new Set<number>();
  doc.descendants((node) => {
    if (node.type.name === "paragraph") {
      const mark = expectParagraphAttrs(node).pPrMark;
      if (mark?.kind === "del" && mark.info.author === author && mark.info.date === date) {
        ids.add(mark.info.id);
      }
    }
    for (const mark of node.marks) {
      if (
        mark.type.name === "deletion" &&
        String(mark.attrs["author"] ?? "") === author &&
        String(mark.attrs["date"] ?? "") === date
      ) {
        const id: unknown = mark.attrs["revisionId"];
        if (typeof id === "number") ids.add(id);
      }
    }
    return true;
  });
  return [...ids];
};

const restoredNote = <TNote extends Footnote | Endnote>(
  note: TNote,
  document: Document,
  reference: Mark,
): TNote => {
  const { styles, theme, numbering } = document.package;
  const doc = footnoteToProseDoc(note.content, {
    ...(styles !== undefined && { styles }),
    ...(theme !== undefined && { theme }),
  });
  const revisions = noteDeletionRevisions(doc, reference);
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
 * The document with the named notes' deletions that went with their
 * reference rejected: the notes whose reference's deletion was rejected get
 * their text back. `before` is the body that still had those deletions.
 */
export const restoreNotes = (
  document: Document,
  keys: Iterable<NoteKey>,
  before: PMNode,
): Document => {
  const deletions = referenceDeletions(before);
  const wanted = new Set([...keys].filter((key) => deletions.has(key)));
  if (wanted.size === 0) return document;
  const restore = <TNote extends Footnote | Endnote>(
    kind: NoteKind,
    notes: TNote[] | undefined,
  ): TNote[] | undefined => {
    if (!notes?.some((note) => wanted.has(noteKey(kind, note.id)))) return notes;
    return notes.map((note) => {
      const reference = deletions.get(noteKey(kind, note.id));
      return reference ? restoredNote(note, document, reference) : note;
    });
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
export const referenceDeletions = (doc: PMNode): Map<NoteKey, Mark> => {
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
  /**
   * `document` with the note changes an earlier {@link reconcile} wrote back
   * and the editor's document has not taken in yet: a save reads the editor
   * right after a write-back, before the host hands the result back.
   */
  withPending: (document: Document) => Document;
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
  // Note changes written back that the editor's document may not hold yet.
  const pending = new Map<NoteKey, { from: Footnote | Endnote; to: Footnote | Endnote }>();
  const recorded =
    (key: NoteKey, edit: NoteEdit): NoteEdit =>
    (note) => {
      const result = edit(note);
      if (result !== note) pending.set(key, { from: note, to: result });
      return result;
    };
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
      const rejected = referenceDeletions(before);
      const edits = new Map<NoteKey, NoteEdit>();
      for (const [key, state] of now) {
        const previous = was.get(key);
        const rejectedDeletion = rejected.get(key);
        if (state === "live" && previous === "deleted" && rejectedDeletion) {
          edits.set(
            key,
            recorded(key, (note) => {
              const restored = restoredNote(note, document, rejectedDeletion);
              if (restored !== note) replaced.set(key, { restored, previous: note });
              return restored;
            }),
          );
        }
        const reference = deletions.get(key);
        if (state === "deleted" && previous === "live" && reference) {
          edits.set(
            key,
            recorded(key, (note) => {
              const earlier = replaced.get(key);
              replaced.delete(key);
              return earlier?.restored === note
                ? earlier.previous
                : deletedNote(note, document, reference);
            }),
          );
        }
      }
      return editNotes(document, edits);
    },
    withPending: (document) => {
      const edits = new Map<NoteKey, NoteEdit>();
      for (const [key, { from, to }] of pending) {
        edits.set(key, (note) => {
          // Taken in (or changed since): nothing is owed any more.
          if (note !== from) pending.delete(key);
          return note === from ? to : note;
        });
      }
      return editNotes(document, edits);
    },
    reset: () => {
      base = null;
      replaced.clear();
      pending.clear();
    },
  };
};

/** The notes model content refers to, live or pending deletion. */
export const noteKeysReferencedIn = (content: unknown): Set<NoteKey> => {
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
  visit(content);
  return keys;
};

/**
 * The notes anything refers to, live or pending deletion: the main story, and
 * the stories the body does not hold (headers, footers, comments, notes).
 * References inside the `skipped` notes do not count: those notes are going.
 */
export const referencedNotes = (
  document: Document,
  skipped: ReadonlySet<string> = new Set(),
): Set<NoteKey> => {
  const { headers, footers, footnotes, endnotes } = document.package;
  const kept = (kind: NoteKind) => (note: Footnote | Endnote) =>
    !skipped.has(noteKey(kind, note.id));
  return noteKeysReferencedIn([
    document.package.document.content,
    [...(headers?.values() ?? []), ...(footers?.values() ?? [])],
    document.package.document.comments ?? [],
    (footnotes ?? []).filter(kept("footnote")),
    (endnotes ?? []).filter(kept("endnote")),
  ]);
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
    (kind === "footnote"
      ? isSeparatorFootnote(note as Footnote)
      : isSeparatorEndnote(note as Endnote)) || referenced.has(noteKey(kind, note.id));
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
