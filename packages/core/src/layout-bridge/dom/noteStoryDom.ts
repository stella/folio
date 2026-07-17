/** Shared painted-DOM identity resolver for editable note stories. */

import type { NoteStoryKey } from "../../controller/noteEditorManager";

type NoteTarget = {
  dataset: DOMStringMap;
};

type ClosestNoteTarget = {
  closest: (selector: string) => NoteTarget | null;
};

const isClosestNoteTarget = (value: unknown): value is ClosestNoteTarget =>
  typeof value === "object" &&
  value !== null &&
  "closest" in value &&
  typeof value.closest === "function";

export const findNoteStoryForTarget = (target: unknown): NoteStoryKey | null => {
  if (!isClosestNoteTarget(target)) return null;
  const noteTarget = target.closest("[data-note-kind][data-note-id]");
  const kind = noteTarget?.dataset["noteKind"];
  const rawNoteId = noteTarget?.dataset["noteId"] ?? "";
  if (!/^-?\d+$/u.test(rawNoteId)) return null;
  const noteId = Number(rawNoteId);
  if ((kind !== "footnote" && kind !== "endnote") || !Number.isSafeInteger(noteId)) return null;
  return { kind, noteId };
};
