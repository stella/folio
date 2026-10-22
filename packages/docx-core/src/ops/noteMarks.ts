/** Automatic note marks are owned model records, derived at the operations boundary. */
import type { BlockContent, Document, Footnote, Endnote, Run } from "../model/document";
import type { NoteStory } from "./types";
import { storyParagraphs, updateBlockList } from "./blocks";
import { leafSpans } from "./leaves";
import { documentStories, storyBody } from "./stories";

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
      leafSpans(paragraph.content).some(({ node }) => node.type === "noteMarker"),
    )
  )
    return note.content;
  const first = paragraphs.at(0);
  if (!first) return note.content;
  const marker = {
    type: "run",
    formatting: { styleId: note.type === "footnote" ? "FootnoteReference" : "EndnoteReference" },
    content: [{ type: "noteMarker", kind: note.type }],
  } as const satisfies Run;
  return updateBlockList(note.content, first.list, (blocks) =>
    blocks.map((block, index) =>
      index === first.index
        ? Object.assign({}, first.paragraph, { content: [marker, ...first.paragraph.content] })
        : block,
    ),
  );
};

export const noteUsesCustomMark = (document: Document, note: Footnote | Endnote): boolean =>
  documentStories(document).some((story) =>
    storyParagraphs(storyBody(document, story)).some(({ paragraph }) =>
      leafSpans(paragraph.content).some(
        ({ node }) =>
          ((note.type === "footnote" && node.type === "footnoteRef") ||
            (note.type === "endnote" && node.type === "endnoteRef")) &&
          node.id === note.id &&
          node.customMarkFollows === true,
      ),
    ),
  );

const referencedNoteKeys = (document: Document): Set<string> => {
  const references = new Set<string>();
  for (const story of documentStories(document)) {
    for (const { paragraph } of storyParagraphs(storyBody(document, story))) {
      for (const { node } of leafSpans(paragraph.content)) {
        if (node.type === "footnoteRef") references.add(`footnote:${String(node.id)}`);
        if (node.type === "endnoteRef") references.add(`endnote:${String(node.id)}`);
      }
    }
  }
  return references;
};

/** Note stories whose final live reference an edit removes. */
export const removedNoteStories = (before: Document, after: Document): readonly NoteStory[] => {
  const beforeReferences = referencedNoteKeys(before);
  const afterReferences = referencedNoteKeys(after);
  return [...beforeReferences]
    .filter((key) => !afterReferences.has(key))
    .map((key) => {
      const [kind, id] = key.split(":");
      return { kind: kind === "footnote" ? "footnote" : "endnote", id: Number(id) };
    });
};

/** Note stories whose ownership changes because an edit adds its first or removes its last reference. */
export const noteStoriesTouchedByReferences = (
  before: Document,
  after: Document,
): readonly NoteStory[] => {
  const beforeReferences = referencedNoteKeys(before);
  const afterReferences = referencedNoteKeys(after);
  return [...new Set([...beforeReferences, ...afterReferences])]
    .filter((key) => beforeReferences.has(key) !== afterReferences.has(key))
    .map((key) => {
      const [kind, id] = key.split(":");
      return { kind: kind === "footnote" ? "footnote" : "endnote", id: Number(id) };
    });
};
