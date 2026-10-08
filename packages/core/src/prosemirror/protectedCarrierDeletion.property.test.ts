import { expect, test } from "bun:test";
import fc from "fast-check";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
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

test(
  "caret deletion owns the complete structured carrier across sizes and revisions",
  async () => {
    await assertProperty(
      fc.property(fc.integer({ min: 1, max: 24 }), (length) => {
        for (const direction of ["Delete", "Backspace"] as const) {
          for (const ownership of ["own", "other", "plain"] as const) {
            for (const protection of ["protected", "plain"] as const) {
              const source = createEmptyDocument();
              const insertion =
                ownership === "plain"
                  ? []
                  : [
                      schema.mark("insertion", {
                        revisionId: 35,
                        author: ownership === "own" ? HARNESS_AUTHOR : "Other Inserter",
                      }),
                    ];
              const child = schema.text(
                "x".repeat(length),
                protection === "protected" ? [protectedRevision] : [],
              );
              const field = schema.node(
                "structuredField",
                {
                  fieldType: "REF",
                  instruction: "REF target",
                  displayText: `before${child.textContent}after`,
                  fieldKind: "simple",
                },
                [schema.text("before"), child, schema.text("after")],
                insertion,
              );
              const doc = schema.node(
                "doc",
                null,
                schema.node("paragraph", null, [schema.text("left"), field, schema.text("right")]),
              );
              const view = new HeadlessEditorView(
                EditorState.create({
                  doc,
                  selection: TextSelection.create(
                    doc,
                    direction === "Delete" ? 5 : 5 + field.nodeSize,
                  ),
                  plugins: createHarnessPlugins(source, "suggesting"),
                }),
              );
              expect(view.pressKey(direction)).toBe(true);
              const retained = view.state.doc.nodeAt(5);
              if (ownership === "own" && protection === "plain") {
                expect(view.state.doc.textContent).toBe("leftright");
              } else {
                expect(retained?.type.name).toBe("structuredField");
                expect(retained?.marks.some(({ type }) => type.name === "deletion")).toBe(true);
                expect(retained?.textContent).toBe(field.textContent);
                if (protection === "protected") {
                  const marks = [];
                  retained?.descendants((node) => {
                    const mark = node.marks.find(
                      ({ attrs, type }) => type.name === "deletion" && attrs["revisionId"] === 37,
                    );
                    if (mark) marks.push(mark.toJSON());
                  });
                  expect(marks).toEqual([protectedRevision.toJSON()]);
                }
              }
              expect(resolveAllChangesInHeadlessState(view.state, "accept").doc.textContent).toBe(
                "leftright",
              );
              const rejectedText =
                ownership === "plain"
                  ? `leftbefore${protection === "plain" ? child.textContent : ""}afterright`
                  : "leftright";
              expect(resolveAllChangesInHeadlessState(view.state, "reject").doc.textContent).toBe(
                rejectedText,
              );
            }
          }
        }
      }),
      { numRuns: 30 },
    );
  },
  propertyTestTimeout(5_000),
);
