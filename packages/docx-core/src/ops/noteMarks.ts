/** Automatic note marks are owned model records, derived at the operations boundary. */
import { panic } from "better-result";

import type {
  BlockContent,
  Document,
  Endnote,
  Footnote,
  Paragraph,
  Run,
} from "../model/document";
import { storyParagraphs, updateBlockList } from "./blocks";
import { type InlineNode, leafSpans, nestedRecords } from "./leaves";
import { documentStories, sameStory, storyBody } from "./stories";
import type { NoteStory } from "./types";

export type NoteReferenceSpan = {
  node: Extract<InlineNode, { type: "footnoteRef" | "endnoteRef" }>;
  ancestors: readonly InlineNode[];
};

/** The automatic marks alone, in their source runs and run properties. */
export const automaticNoteMarkRuns = (nodes: readonly InlineNode[]): Run[] =>
  leafSpans(nodes).flatMap(({ node, ancestors }) => {
    if (node.type !== "noteMarker") return [];
    const run =
      ancestors.findLast((ancestor) => ancestor.type === "run") ??
      panic("An automatic note mark belongs to a source run.");
    return [{ ...run, content: [node] }];
  });

/**
 * Every footnote and endnote reference in `nodes`, with the records holding
 * it, outermost first: the one reading of where a document references a
 * note. A field is one unit of text, but its result is
 * content a reader sees and can hold a reference, so the walk enters it
 * ({@link nestedRecords}), nested fields included. A field's instruction
 * is not content and holds none.
 */
export const noteReferenceSpans = (
  nodes: readonly InlineNode[],
  ancestors: readonly InlineNode[] = [],
): NoteReferenceSpan[] =>
  nodes.flatMap((node): NoteReferenceSpan[] => {
    if (node.type === "footnoteRef" || node.type === "endnoteRef")
      return [{ node, ancestors }];
    const nested = nestedRecords(node);
    return nested === undefined
      ? []
      : noteReferenceSpans(nested, [...ancestors, node]);
  });

type NoteMarkOptions = { note: Footnote | Endnote; customMark: boolean };
export const noteContentWithAutomaticMark = ({
  note,
  customMark,
}: NoteMarkOptions): BlockContent[] => {
  if (customMark || (note.noteType !== undefined && note.noteType !== "normal"))
    return note.content;
  const paragraphs = storyParagraphs({ content: note.content });
  if (
    paragraphs.some(({ paragraph }) =>
      leafSpans(paragraph.content).some(
        ({ node }) => node.type === "noteMarker",
      ),
    )
  )
    return note.content;
  const first = paragraphs.at(0);
  if (!first) return note.content;
  const marker = {
    type: "run",
    formatting: {
      styleId:
        note.type === "footnote" ? "FootnoteReference" : "EndnoteReference",
    },
    content: [{ type: "noteMarker", kind: note.type }],
  } as const satisfies Run;
  return updateBlockList(note.content, first.list, (blocks) =>
    blocks.map((block, index) =>
      index === first.index
        ? Object.assign({}, first.paragraph, {
            content: [marker, ...first.paragraph.content],
          })
        : block,
    ),
  );
};

export const noteUsesCustomMark = (
  document: Document,
  note: Footnote | Endnote,
): boolean =>
  documentStories(document).some((story) =>
    storyParagraphs(storyBody(document, story)).some(({ paragraph }) =>
      noteReferenceSpans(paragraph.content).some(
        ({ node }) =>
          ((note.type === "footnote" && node.type === "footnoteRef") ||
            (note.type === "endnote" && node.type === "endnoteRef")) &&
          node.id === note.id &&
          node.customMarkFollows === true,
      ),
    ),
  );

/** The notes one paragraph references. */
const paragraphNoteKeys = (paragraph: Paragraph): readonly string[] => {
  const keys: string[] = [];
  for (const { node } of noteReferenceSpans(paragraph.content)) {
    if (node.type === "footnoteRef") keys.push(`footnote:${String(node.id)}`);
    if (node.type === "endnoteRef") keys.push(`endnote:${String(node.id)}`);
  }
  return keys;
};

/**
 * The notes a document references, read once per document and per paragraph
 * object. Operations never change a record they are given, so a paragraph an
 * edit left as it was reads as it did, and only the paragraphs an edit
 * replaced are read again.
 */
const PARAGRAPH_NOTE_KEYS = new WeakMap<Paragraph, readonly string[]>();
const DOCUMENT_NOTE_KEYS = new WeakMap<Document, ReadonlySet<string>>();
const sharedParagraphNoteKeys = (paragraph: Paragraph): readonly string[] => {
  const known = PARAGRAPH_NOTE_KEYS.get(paragraph);
  if (known !== undefined) return known;
  const keys = Object.freeze(paragraphNoteKeys(paragraph));
  PARAGRAPH_NOTE_KEYS.set(paragraph, keys);
  return keys;
};

const noteKeysOf =
  (ofParagraph: (paragraph: Paragraph) => readonly string[]) =>
  (document: Document): Set<string> => {
    const references = new Set<string>();
    for (const story of documentStories(document))
      for (const { paragraph } of storyParagraphs(storyBody(document, story)))
        for (const key of ofParagraph(paragraph)) references.add(key);
    return references;
  };

/** {@link referencedNoteKeys} reading every paragraph again: its reference. */
export const referencedNoteKeysInFull: (document: Document) => Set<string> =
  noteKeysOf(paragraphNoteKeys);

const readSharedNoteKeys = noteKeysOf(sharedParagraphNoteKeys);
export const referencedNoteKeys = (document: Document): ReadonlySet<string> => {
  const known = DOCUMENT_NOTE_KEYS.get(document);
  if (known !== undefined) return known;
  const keys = readSharedNoteKeys(document);
  DOCUMENT_NOTE_KEYS.set(document, keys);
  return keys;
};

/** The note keys a run of blocks references, sorted. */
const noteKeysIn = (blocks: readonly BlockContent[]): string[] =>
  storyParagraphs({ content: [...blocks] })
    .flatMap(({ paragraph }) => sharedParagraphNoteKeys(paragraph))
    .sort();

/**
 * Whether `after` references exactly the notes `before` does, read from the
 * top-level blocks of each story that are not the very blocks `before` held
 * at the same place from either end: when those carry the same references
 * as the ones they replace, every other block's references are unchanged.
 * False when the stories differ or the references might.
 */
export const sameNoteReferences = (
  before: Document,
  after: Document,
): boolean => {
  const stories = documentStories(after);
  const earlier = documentStories(before);
  if (
    stories.length !== earlier.length ||
    stories.some((story, index) => {
      const other = earlier[index];
      return other === undefined || !sameStory(story, other);
    })
  )
    return false;
  for (const story of stories) {
    const old = storyBody(before, story).content;
    const now = storyBody(after, story).content;
    if (old === now) continue;
    let start = 0;
    while (
      start < old.length &&
      start < now.length &&
      old[start] === now[start]
    )
      start += 1;
    let end = 0;
    while (
      end < old.length - start &&
      end < now.length - start &&
      old[old.length - 1 - end] === now[now.length - 1 - end]
    )
      end += 1;
    const removed = noteKeysIn(old.slice(start, old.length - end));
    const added = noteKeysIn(now.slice(start, now.length - end));
    if (
      removed.length !== added.length ||
      removed.some((key, index) => added[index] !== key)
    )
      return false;
  }
  return true;
};

/** Note stories whose final live reference an edit removes. */
export const removedNoteStories = (
  before: Document,
  after: Document,
): readonly NoteStory[] => {
  if (sameNoteReferences(before, after)) return [];
  const beforeReferences = referencedNoteKeys(before);
  const afterReferences = referencedNoteKeys(after);
  return [...beforeReferences]
    .filter((key) => !afterReferences.has(key))
    .map((key) => {
      const [kind, id] = key.split(":");
      return {
        kind: kind === "footnote" ? "footnote" : "endnote",
        id: Number(id),
      };
    });
};

/** Note stories whose ownership changes because an edit adds its first or removes its last reference. */
export const noteStoriesTouchedByReferences = (
  before: Document,
  after: Document,
): readonly NoteStory[] => {
  if (sameNoteReferences(before, after)) return [];
  const beforeReferences = referencedNoteKeys(before);
  const afterReferences = referencedNoteKeys(after);
  return [...new Set([...beforeReferences, ...afterReferences])]
    .filter((key) => beforeReferences.has(key) !== afterReferences.has(key))
    .map((key) => {
      const [kind, id] = key.split(":");
      return {
        kind: kind === "footnote" ? "footnote" : "endnote",
        id: Number(id),
      };
    });
};
