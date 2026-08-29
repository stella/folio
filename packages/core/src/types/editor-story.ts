/** A secondary note story address, independent of any mounted editor view. */
export type NoteStoryKey = {
  kind: "footnote" | "endnote";
  noteId: number;
};
