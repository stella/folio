import { describe, expect, test } from "bun:test";
import type { EditorView } from "prosemirror-view";

import { resolveActiveEditorStory } from "./activeEditorStory";

/* oxlint-disable typescript/no-unsafe-type-assertion -- opaque view sentinels */
const bodyView = { story: "body" } as unknown as EditorView;
const headerFooterView = { story: "headerFooter" } as unknown as EditorView;
const noteView = { story: "note" } as unknown as EditorView;
/* oxlint-enable typescript/no-unsafe-type-assertion */

describe("resolveActiveEditorStory", () => {
  test("uses note, header/footer, then body precedence", () => {
    expect(resolveActiveEditorStory({ bodyView, headerFooterView, noteView })).toEqual({
      type: "note",
      view: noteView,
    });
    expect(resolveActiveEditorStory({ bodyView, headerFooterView, noteView: null })).toEqual({
      type: "headerFooter",
      view: headerFooterView,
    });
    expect(resolveActiveEditorStory({ bodyView, headerFooterView: null, noteView: null })).toEqual({
      type: "body",
      view: bodyView,
    });
  });

  test("returns an explicit empty target before any view mounts", () => {
    expect(
      resolveActiveEditorStory({ bodyView: null, headerFooterView: null, noteView: null }),
    ).toEqual({ type: "none", view: null });
  });
});
