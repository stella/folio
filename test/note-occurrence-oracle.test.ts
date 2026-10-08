import { expect, test } from "bun:test";
import { normalizeNoteOccurrenceIds } from "./note-occurrence-oracle";

test("independent occurrence identities compare by equivalence class without hiding aliases", () => {
  const fragments = (ids: string[]) => ids.map((occurrenceId) => ({ occurrenceId, text: "123" }));
  expect(normalizeNoteOccurrenceIds(fragments(["a", "a", "b"]))).toEqual(
    normalizeNoteOccurrenceIds(fragments(["c", "c", "d"])),
  );
  expect(normalizeNoteOccurrenceIds(fragments(["a", "a", "b"]))).not.toEqual(
    normalizeNoteOccurrenceIds(fragments(["c", "c", "c"])),
  );
});
