import { describe, expect, test } from "bun:test";

import { acquireHarnessDom } from "../../../../test/harnessDom";
import type { Document } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import {
  createNoteEditorManager,
  enumerateDocumentNoteStories,
} from "./noteEditorManager";

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
          {
            type: "endnote",
            id: 0,
            noteType: "continuationSeparator",
            content: [],
          },
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

describe("note editor sync", () => {
  test("reads the document only while a note story is mounted or active", () => {
    const releaseDom = acquireHarnessDom();
    const host = document.createElement("div");
    document.body.append(host);
    const source = createEmptyDocument({ initialText: "Body" });
    source.package.footnotes = [
      {
        type: "footnote",
        id: 1,
        content: createEmptyDocument({ initialText: "Note" }).package.document
          .content,
      },
    ];
    let reads = 0;
    const manager = createNoteEditorManager({
      getHost: () => host,
      getDocument: () => {
        reads += 1;
        return source;
      },
      getStyles: () => null,
      getTheme: () => null,
    });
    try {
      manager.sync();
      manager.sync();
      expect(reads).toBe(0);
      expect(manager.getActive()).toBeNull();

      const view = manager.activate({ kind: "footnote", noteId: 1 });
      expect(view?.state.doc.textContent).toBe("Note");
      const afterActivate = reads;
      expect(afterActivate).toBeGreaterThan(0);

      // A story stays mounted after it stops being active, and keeps syncing.
      manager.activate(null);
      const afterDeactivate = reads;
      manager.sync();
      expect(reads).toBeGreaterThan(afterDeactivate);
      expect(manager.getView({ kind: "footnote", noteId: 1 })).toBe(view);
    } finally {
      manager.destroy();
      host.remove();
      releaseDom();
    }
  });
});
