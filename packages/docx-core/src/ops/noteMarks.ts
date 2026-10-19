/** Automatic note marks are owned model records, derived at the operations boundary. */
import type { BlockContent, Document, Footnote, Endnote, Run } from "../model/document";
import { storyParagraphs, updateBlockList } from "./blocks";
import { leafSpans } from "./leaves";
import { documentStories, replaceStoryBody, storyBody } from "./stories";

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

/** Restore source-owned automatic markers after an operation rewrites a note story. */
export const withAutomaticNoteMarks = (document: Document): Document => {
  let current = document;
  for (const story of documentStories(document)) {
    if (story === "main" || (story.kind !== "footnote" && story.kind !== "endnote")) continue;
    const note =
      story.kind === "footnote"
        ? current.package.footnotes?.find(({ id }) => id === story.id)
        : current.package.endnotes?.find(({ id }) => id === story.id);
    if (!note) continue;
    const content = noteContentWithAutomaticMark({
      note,
      customMark: noteUsesCustomMark(current, note),
    });
    if (content === note.content) continue;
    current = replaceStoryBody({ document: current, story, body: { content } });
  }
  return current;
};
