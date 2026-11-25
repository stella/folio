/** Read note references and retain the corresponding model stories. */
import type { Node as PMNode } from "prosemirror-model";

import { setDocumentBodyComments } from "@stll/docx-core/model";

import { withoutLostCommentThreads } from "../ai-edits/comment-lifecycle";
import { visitCommentMarkers } from "../docx/commentAnchorIndex";
import {
  isSeparatorEndnote,
  isSeparatorFootnote,
} from "../docx/footnoteParser";
import type {
  BlockContent,
  Document,
  Endnote,
  Footnote,
} from "../types/document";
import { expectFootnoteRefMarkAttrs } from "./attrs";

type NoteKind = "footnote" | "endnote";

/** `footnote:<id>` or `endnote:<id>`. */
export type NoteKey = `${NoteKind}:${number}`;

const noteKey = (kind: NoteKind, id: number): NoteKey => `${kind}:${id}`;

/**
 * The notes a story's text refers to, and whether each reference is pending
 * deletion: a reference is deleted when every character of it is.
 */
export const noteReferenceStates = (
  doc: PMNode,
): Map<NoteKey, "live" | "deleted"> => {
  const states = new Map<NoteKey, "live" | "deleted">();
  doc.descendants((node) => {
    if (!node.isText) return true;
    const reference = node.marks.find(
      (mark) => mark.type.name === "footnoteRef",
    );
    if (!reference) return false;
    const { id, noteType } = expectFootnoteRefMarkAttrs(reference);
    const key = noteKey(
      noteType === "endnote" ? "endnote" : "footnote",
      Number(id),
    );
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
    if (
      (type === "footnoteRef" || type === "endnoteRef") &&
      typeof id === "number"
    ) {
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
const anchoredComments = (
  stories: readonly (readonly BlockContent[])[],
): Set<number> => {
  const ids = new Set<number>();
  for (const blocks of stories)
    visitCommentMarkers(blocks, ({ item }) => ids.add(item.id));
  return ids;
};

/**
 * The notes `source` holds that nothing in it refers to: the package's own
 * content, which no edit has touched. Taken when a document is opened, so a
 * save can tell them from the notes an edit left without a reference.
 */
export const heldUnreferencedNotes = (
  source: Document,
): ReadonlySet<string> => {
  const referenced = referencedNotes(source);
  const held = new Set<NoteKey>();
  const { footnotes, endnotes } = source.package;
  for (const note of footnotes ?? []) {
    const key = noteKey("footnote", note.id);
    if (!referenced.has(key)) held.add(key);
  }
  for (const note of endnotes ?? []) {
    const key = noteKey("endnote", note.id);
    if (!referenced.has(key)) held.add(key);
  }
  return held;
};

/**
 * The document without the notes an edit left with no reference: a note goes
 * with its reference, so accepting the reference's deletion leaves a note no
 * reader can reach, and saving drops it, with the comments anchored only in
 * it. A note the opened package already held without a reference
 * (`heldUnreferenced`, from {@link heldUnreferencedNotes}) is the package's
 * content, not an edit's leftover, and stays. Separators stay.
 */
export const withoutUnreferencedNotes = (
  document: Document,
  heldUnreferenced: ReadonlySet<string>,
): Document => {
  const referenced = referencedNotes(document);
  const { footnotes, endnotes, headers, footers } = document.package;
  const keeps = (kind: NoteKind) => (note: Footnote | Endnote) =>
    (kind === "footnote"
      ? isSeparatorFootnote(note as Footnote)
      : isSeparatorEndnote(note as Endnote)) ||
    referenced.has(noteKey(kind, note.id)) ||
    heldUnreferenced.has(noteKey(kind, note.id));
  const keptFootnotes = footnotes?.filter(keeps("footnote"));
  const keptEndnotes = endnotes?.filter(keeps("endnote"));
  if (
    keptFootnotes?.length === footnotes?.length &&
    keptEndnotes?.length === endnotes?.length
  ) {
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
          ...[...(keptFootnotes ?? []), ...(keptEndnotes ?? [])].map(
            (note) => note.content,
          ),
        ]),
      })
    : undefined;
  const documentBody = { ...document.package.document };
  if (keptComments !== undefined && keptComments.length !== comments?.length) {
    setDocumentBodyComments(documentBody, keptComments);
  }
  return {
    ...document,
    package: {
      ...document.package,
      ...(keptFootnotes !== undefined && { footnotes: keptFootnotes }),
      ...(keptEndnotes !== undefined && { endnotes: keptEndnotes }),
      ...(keptComments !== undefined &&
        keptComments.length !== comments?.length && { document: documentBody }),
    },
  };
};
