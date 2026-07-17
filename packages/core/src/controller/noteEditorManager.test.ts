import { describe, expect, test } from "bun:test";

import type { Document } from "../types/document";
import { enumerateDocumentNoteStories } from "./noteEditorManager";

describe("note editor story identity", () => {
  test("keeps footnote and endnote ids in separate namespaces", () => {
    const document: Document = {
      package: {
        document: { content: [] },
        footnotes: [{ type: "footnote", id: 4, content: [] }],
        endnotes: [{ type: "endnote", id: 4, content: [] }],
      },
    };

    expect(enumerateDocumentNoteStories(document)).toEqual([
      { kind: "footnote", noteId: 4 },
      { kind: "endnote", noteId: 4 },
    ]);
  });

  test("does not expose separator stories as editable note bodies", () => {
    const document: Document = {
      package: {
        document: { content: [] },
        footnotes: [
          { type: "footnote", id: -1, noteType: "separator", content: [] },
          { type: "footnote", id: 2, noteType: "normal", content: [] },
        ],
        endnotes: [
          { type: "endnote", id: 0, noteType: "continuationSeparator", content: [] },
          { type: "endnote", id: 3, content: [] },
        ],
      },
    };

    expect(enumerateDocumentNoteStories(document)).toEqual([
      { kind: "footnote", noteId: 2 },
      { kind: "endnote", noteId: 3 },
    ]);
  });
});
