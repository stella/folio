import { describe, expect, test } from "bun:test";

import { findNoteStoryForTarget } from "./noteStoryDom";

const target = (dataset: Record<string, string>) => ({
  closest: () => ({ dataset }),
});

describe("painted note story identity", () => {
  test("resolves both note namespaces from a reference or body ancestor", () => {
    expect(findNoteStoryForTarget(target({ noteKind: "footnote", noteId: "12" }))).toEqual({
      kind: "footnote",
      noteId: 12,
    });
    expect(findNoteStoryForTarget(target({ noteKind: "endnote", noteId: "12" }))).toEqual({
      kind: "endnote",
      noteId: 12,
    });
  });

  test("rejects incomplete and malformed story markers", () => {
    expect(findNoteStoryForTarget(target({ noteKind: "footnote" }))).toBeNull();
    expect(findNoteStoryForTarget(target({ noteKind: "comment", noteId: "2" }))).toBeNull();
    expect(findNoteStoryForTarget(target({ noteKind: "footnote", noteId: "2x" }))).toBeNull();
    expect(findNoteStoryForTarget(null)).toBeNull();
  });
});
