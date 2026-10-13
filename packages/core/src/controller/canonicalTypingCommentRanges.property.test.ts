import { expect, test } from "bun:test";
import fc from "fast-check";
import { EditorState, TextSelection } from "prosemirror-state";
import { paragraphLogicalText } from "@stll/docx-core/ops";
import type { Document, ParagraphContent } from "../types/document";
import { getTrackedChangesFromDoc } from "../ai-edits/read";
import { assertProperty } from "../../../../test/property-testing";
import { createCanonicalSession, publishCanonicalProjection } from "./canonicalSession";

// Projection fixtures edited at fixed positions; none followed the caret after replacement.
test("successive typing follows the replacement caret across comment boundaries", () => {
  assertProperty(
    fc.property(
      fc.constantFrom("x", "é", "👍🏽"),
      fc.integer({ min: 0, max: 3 }),
      (first, references) => {
        for (const type of ["editing", "suggesting"] as const) {
          const atoms = Array.from(
            { length: references },
            (_, id) =>
              ({
                type: "run",
                content: [{ type: "commentReference", id }],
              }) satisfies ParagraphContent,
          );
          const source = {
            package: {
              document: {
                content: [
                  {
                    type: "paragraph",
                    paraId: "12345678",
                    content: [
                      { type: "commentRangeStart", id: 0 },
                      { type: "run", content: [{ type: "text", text: "Commented" }] },
                      { type: "commentRangeEnd", id: 0 },
                      ...atoms,
                      { type: "run", content: [{ type: "text", text: " tail" }] },
                    ],
                  },
                ],
              },
            },
          } satisfies Document;
          const session = createCanonicalSession(source).unwrap();
          session.setMode(type === "editing" ? { type } : { type, author: "Reviewer" });
          let state = EditorState.create({ doc: session.projection.doc });
          state = state.apply(
            state.tr.setSelection(TextSelection.create(state.doc, 1, state.doc.content.size - 1)),
          );
          const typed = [first, "y", "z"];
          for (const text of typed) {
            const { from, to } = state.selection;
            const commit = session
              .prepareReplace(state, { from, to, text, semantic: "typing" })
              .unwrap();
            state = publishCanonicalProjection({ session, state, commit }).unwrap().state;
          }
          const revisionIds = [...new Set(getTrackedChangesFromDoc(state.doc).map(({ id }) => id))];
          if (revisionIds.length > 0) {
            const commit = session
              .prepareResolve(state, { revisionIds, resolution: "accept" })
              .unwrap();
            state = publishCanonicalProjection({ session, state, commit }).unwrap().state;
          }
          const paragraph = session.document.package.document.content.at(0);
          if (paragraph?.type !== "paragraph") throw new Error("Expected paragraph");
          expect(paragraphLogicalText(paragraph)).toBe(typed.join(""));
        }
      },
    ),
    { seed: 1404001, numRuns: 12 },
  );
});
