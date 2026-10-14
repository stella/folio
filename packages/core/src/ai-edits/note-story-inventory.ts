import type { Document } from "../types/document";

type NoteStories = Pick<Document["package"], "footnotes" | "endnotes">;

/** The note collections visible to a reviewer or its captured save state. */
export const noteStoryInventory = (
  source: NoteStories,
  removed: Pick<ReadonlySet<string>, "has" | "size">,
): NoteStories => {
  if (removed.size === 0) return source;
  return {
    ...(source.footnotes && {
      footnotes: source.footnotes.filter((note) => !removed.has(`footnote:${note.id}`)),
    }),
    ...(source.endnotes && {
      endnotes: source.endnotes.filter((note) => !removed.has(`endnote:${note.id}`)),
    }),
  };
};
