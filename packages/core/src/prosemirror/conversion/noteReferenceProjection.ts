/**
 * What the body's note references say about the notes the package keeps.
 *
 * A note belongs to its reference. When every reference to a note is a
 * tracked deletion, the note's content is deleted with it — its runs and its
 * paragraph marks, the in-note reference mark included — so accepting the
 * deletion takes the note away and rejecting it keeps the note. A note no
 * reference points to any more (deleted outright, a deletion accepted, or a
 * pending insertion taken back) goes too instead of staying behind
 * unreferenced, unless another story — a header, a footer, a comment, another
 * note — still refers to it. That holds for a note the document's body
 * referenced when it was read and for a note made during the session; a note
 * the package held without any reference is left as it was.
 *
 * The body is the only place a reference lives, so the notes are projected
 * from it at save rather than kept in step edit by edit.
 */

import type { Node as PMNode } from "prosemirror-model";

import { isSessionNote, withNoteReferenceMark } from "../../docx/noteReferenceMark";

import type {
  BlockContent,
  Document,
  Endnote,
  Footnote,
  Paragraph,
  ParagraphContent,
  TrackedChangeInfo,
} from "../../types/document";

type NoteKind = "footnote" | "endnote";

type ReferenceState = {
  /** A reference that is not a tracked deletion. */
  live: boolean;
  /** The tracked deletion of the first deleted reference. */
  deletion: TrackedChangeInfo | null;
};

const key = (kind: NoteKind, id: number | string): string => `${kind}:${String(id)}`;

const numberAttr = (value: unknown): number | undefined =>
  typeof value === "number" && Number.isFinite(value) ? value : undefined;

const stringAttr = (value: unknown): string | undefined =>
  typeof value === "string" ? value : undefined;

/** A note reference spelled in markup the body keeps verbatim (a drawing's text, say). */
const VERBATIM_REFERENCE =
  /<(?:[\w.-]+:)?(?<kind>footnote|endnote)Reference\b[^>]*?\bw:id="(?<id>-?\d+)"/gu;

/** The state of every note reference in the ProseMirror body. */
const bodyReferenceStates = (doc: PMNode): Map<string, ReferenceState> => {
  const states = new Map<string, ReferenceState>();
  doc.descendants((node) => {
    if (!node.isText && node.isLeaf && node.type.name !== "image") {
      // Markup kept verbatim still refers to its notes, and cannot be edited.
      const markup = JSON.stringify(node.attrs);
      if (markup.includes("Reference")) {
        for (const match of markup.replaceAll('\\"', '"').matchAll(VERBATIM_REFERENCE)) {
          const kind: NoteKind = match.groups?.["kind"] === "endnote" ? "endnote" : "footnote";
          const id = key(kind, match.groups?.["id"] ?? "");
          states.set(id, { deletion: null, ...states.get(id), live: true });
        }
      }
    }
    const reference = node.isText
      ? node.marks.find((mark) => mark.type.name === "footnoteRef")
      : undefined;
    if (!reference) {
      return true;
    }
    const kind: NoteKind = reference.attrs["noteType"] === "endnote" ? "endnote" : "footnote";
    const id = key(kind, reference.attrs["id"] as number | string);
    const deletionMark = node.marks.find(
      (mark) => mark.type.name === "deletion" && mark.attrs["provenance"] !== "suggested",
    );
    const state = states.get(id) ?? { live: false, deletion: null };
    if (!deletionMark) {
      state.live = true;
    } else if (!state.deletion) {
      const revisionId = numberAttr(deletionMark.attrs["revisionId"]);
      const author = stringAttr(deletionMark.attrs["author"]);
      const date = stringAttr(deletionMark.attrs["date"]);
      if (revisionId !== undefined && author !== undefined) {
        state.deletion = { id: revisionId, author, ...(date ? { date } : {}) };
      }
    }
    states.set(id, state);
    return true;
  });
  return states;
};

/** Every note the given model content references, found wherever a run holds one. */
const modelReferences = (content: unknown): Set<string> => {
  const found = new Set<string>();
  const visit = (value: unknown): void => {
    if (Array.isArray(value)) {
      for (const item of value) {
        visit(item);
      }
      return;
    }
    if (typeof value !== "object" || value === null) {
      return;
    }
    const record = value as Record<string, unknown>;
    if (
      (record["type"] === "footnoteRef" || record["type"] === "endnoteRef") &&
      (typeof record["id"] === "number" || typeof record["id"] === "string")
    ) {
      found.add(key(record["type"] === "endnoteRef" ? "endnote" : "footnote", record["id"]));
    }
    for (const entry of Object.values(record)) {
      visit(entry);
    }
  };
  visit(content);
  return found;
};

const deletedParagraph = (paragraph: Paragraph, info: TrackedChangeInfo): Paragraph => {
  const content: ParagraphContent[] = [];
  for (const item of paragraph.content) {
    switch (item.type) {
      case "run":
      case "hyperlink":
      case "simpleField":
      case "complexField":
      case "inlineSdt":
        content.push({ type: "deletion", info, content: [item] });
        break;
      default:
        content.push(item);
    }
  }
  return {
    ...paragraph,
    content,
    ...(paragraph.pPrMark ? {} : { pPrMark: { kind: "del", info } }),
  };
};

const deletedContent = (blocks: readonly BlockContent[], info: TrackedChangeInfo): BlockContent[] =>
  blocks.map((block) => (block.type === "paragraph" ? deletedParagraph(block, info) : block));

/**
 * The note each projected (deleted) note was made from. An editor saves again
 * against what its last save produced, so rejecting the reference deletion
 * has to find the note as it was before its content was deleted.
 */
const projectionSources = new WeakMap<Footnote | Endnote, Footnote | Endnote>();

const projectNotes = <Note extends Footnote | Endnote>(
  notes: readonly Note[] | undefined,
  kind: NoteKind,
  states: ReadonlyMap<string, ReferenceState>,
  referencedBefore: ReadonlySet<string>,
  referencedElsewhere: ReadonlySet<string>,
): Note[] | undefined => {
  if (!notes) {
    return undefined;
  }
  const projected: Note[] = [];
  for (const note of notes) {
    if (note.noteType !== undefined && note.noteType !== "normal") {
      projected.push(note);
      continue;
    }
    const state = states.get(key(kind, note.id));
    if (!state) {
      // Nothing in the body points to it any more.
      const source = (projectionSources.get(note) as Note | undefined) ?? note;
      const belongedToAReference =
        referencedBefore.has(key(kind, note.id)) || isSessionNote(source);
      if (!belongedToAReference || referencedElsewhere.has(key(kind, note.id))) {
        projected.push(note);
      }
      continue;
    }
    // A note deleted with its reference at an earlier save of the same
    // session comes back as that save left it: start from what it was.
    const source = (projectionSources.get(note) as Note | undefined) ?? note;
    if (!state.live && state.deletion) {
      // The in-note reference mark is part of what the deletion takes.
      const deleted = {
        ...source,
        content: deletedContent(withNoteReferenceMark(kind, source.content), state.deletion),
      };
      projectionSources.set(deleted, source);
      projected.push(deleted);
    } else {
      projected.push(source);
    }
  }
  return projected;
};

/** The package's footnotes and endnotes as the body's references leave them. */
export const projectNotesFromReferences = (
  doc: PMNode,
  base: Document,
): Pick<Document["package"], "footnotes" | "endnotes"> => {
  const { footnotes, endnotes } = base.package;
  if (!footnotes && !endnotes) {
    return {};
  }
  const states = bodyReferenceStates(doc);
  // The stories other than the body, which the editor's body does not hold.
  const referencedElsewhere = modelReferences([
    ...(base.package.headers?.values() ?? []),
    ...(base.package.footers?.values() ?? []),
    base.package.document.comments ?? [],
    footnotes ?? [],
    endnotes ?? [],
  ]);
  const referencedBefore = modelReferences(base.package.document.content);
  const projectedFootnotes = projectNotes(
    footnotes,
    "footnote",
    states,
    referencedBefore,
    referencedElsewhere,
  );
  const projectedEndnotes = projectNotes(
    endnotes,
    "endnote",
    states,
    referencedBefore,
    referencedElsewhere,
  );
  return {
    ...(projectedFootnotes ? { footnotes: projectedFootnotes } : {}),
    ...(projectedEndnotes ? { endnotes: projectedEndnotes } : {}),
  };
};
