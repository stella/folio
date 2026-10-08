import { expect, test } from "bun:test";
import { EditorState, TextSelection } from "prosemirror-state";

import {
  HARNESS_AUTHOR,
  HeadlessEditorView,
  createHarnessPlugins,
} from "../__tests__/editorHarness";
import { createDocx } from "../docx/rezip";
import { parseDocx } from "../docx/parser";
import { createEmptyDocument } from "../utils/createDocument";
import { expectTrackedChangeMarkAttrs } from "./attrs";
import { resolveAllChangesInHeadlessState } from "./commands/comments";
import { fromProseDoc } from "./conversion/fromProseDoc";
import { toProseDoc } from "./conversion/toProseDoc";
import { schema } from "./schema/index";

const protectedRevision = schema.mark("deletion", {
  revisionId: 37,
  author: "Other Reviewer",
  date: "2026-01-02T03:04:05Z",
  initials: "OR",
  _docxRevisionAncestors: [
    { type: "insertion", revisionId: 36, author: "First Reviewer", outerWrapperCount: 0 },
  ],
});

for (const direction of ["Delete", "Backspace"] as const) {
  test(`${direction} preserves another author's deletion inside an own-inserted field`, async () => {
    const source = createEmptyDocument();
    const hyperlink = schema.mark("hyperlink", { href: "https://example.test/field" });
    const field = schema.node(
      "structuredField",
      {
        fieldType: "REF",
        instruction: "REF target",
        displayText: "beforehiddenafter",
        fieldKind: "simple",
      },
      [
        schema.text("before", [hyperlink]),
        schema.text("hidden", [hyperlink, protectedRevision]),
        schema.text("after", [hyperlink]),
      ],
      [schema.mark("insertion", { revisionId: 35, author: HARNESS_AUTHOR })],
    );
    const doc = schema.node(
      "doc",
      null,
      schema.node("paragraph", null, [schema.text("left"), field, schema.text("right")]),
    );
    const fieldPosition = 5;
    const caret = direction === "Delete" ? fieldPosition : fieldPosition + field.nodeSize;
    const view = new HeadlessEditorView(
      EditorState.create({
        doc,
        selection: TextSelection.create(doc, caret),
        plugins: createHarnessPlugins(source, "suggesting"),
      }),
    );
    expect(view.pressKey(direction)).toBe(true);
    const revisions = (state: EditorState) => {
      const found = [];
      state.doc.descendants((node) => {
        for (const mark of node.marks) {
          if (mark.type.name !== "deletion") continue;
          const attrs = expectTrackedChangeMarkAttrs(mark);
          if (attrs.revisionId === 37) found.push(mark.toJSON());
        }
      });
      return found;
    };
    const expected = [protectedRevision.toJSON()];
    expect(revisions(view.state)).toEqual(expected);
    for (const mode of ["accept", "reject"] as const) {
      const live = resolveAllChangesInHeadlessState(view.state, mode);
      expect(live.doc.textContent).toBe("leftright");
      // Resolve before saving: pending child revisions in simple fields are not
      // representable by the current DOCX field model.
      const saved = await createDocx(fromProseDoc(live.doc, source));
      const reopened = toProseDoc(await parseDocx(saved, { preloadFonts: false }));
      expect(reopened.textContent).toBe(live.doc.textContent);
    }
  });
}
