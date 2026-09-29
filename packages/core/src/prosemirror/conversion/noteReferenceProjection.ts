/**
 * What the body's note references say about the notes the package keeps.
 *
 * A note belongs to its reference. When every reference to a note is a
 * tracked deletion, the note's content is deleted with it — its runs and its
 * paragraph marks, the in-note reference mark included — so accepting the
 * deletion takes the note away and rejecting it keeps the note. When the
 * references the body had are gone altogether (deleted outright, or a
 * deletion accepted), the note goes too instead of staying behind unreferenced.
 *
 * The body is the only place a reference lives, so the notes are projected
 * from it at save rather than kept in step edit by edit.
 */

import type { Node as PMNode } from "prosemirror-model";

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

/** The state of every note reference in the ProseMirror body. */
const bodyReferenceStates = (doc: PMNode): Map<string, ReferenceState> => {
  const states = new Map<string, ReferenceState>();
  doc.descendants((node) => {
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

/** Every note the model body references, found wherever a run holds one. */
const modelReferences = (content: readonly BlockContent[]): Set<string> => {
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
      // Referenced when the document was read, and by nothing now.
      if (!referencedBefore.has(key(kind, note.id))) {
        projected.push(note);
      }
      continue;
    }
    // A note deleted with its reference at an earlier save of the same
    // session comes back as that save left it: start from what it was.
    const source = (projectionSources.get(note) as Note | undefined) ?? note;
    if (!state.live && state.deletion) {
      const deleted = { ...source, content: deletedContent(source.content, state.deletion) };
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
  const referencedBefore = modelReferences(base.package.document.content);
  const projectedFootnotes = projectNotes(footnotes, "footnote", states, referencedBefore);
  const projectedEndnotes = projectNotes(endnotes, "endnote", states, referencedBefore);
  return {
    ...(projectedFootnotes ? { footnotes: projectedFootnotes } : {}),
    ...(projectedEndnotes ? { endnotes: projectedEndnotes } : {}),
  };
};
