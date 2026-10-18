import { expect, test } from "bun:test";
import fc from "fast-check";
import { EditorState, TextSelection } from "prosemirror-state";
import type { Document, ParagraphContent, Comment } from "../types/document";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { getTrackedChangeStatsFromDoc } from "../ai-edits/read";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { createCanonicalSession, publishCanonicalProjection } from "./canonicalSession";

// Projection fixtures edited at fixed positions; none followed the caret after replacement.
test(
  "successive typing follows the replacement caret across comment boundaries",
  () => {
    assertProperty(
      fc.property(
        fc.constantFrom("x", "é", "👍🏽"),
        fc.integer({ min: 0, max: 3 }),
        (first, references) => {
          const coveredText = new Map<string, string>();
          for (const type of ["editing", "suggesting"] as const) {
            const atoms = Array.from(
              { length: references },
              (_, id) =>
                ({
                  type: "commentReference",
                  id,
                }) satisfies ParagraphContent,
            );
            const source = {
              package: {
                document: {
                  comments: Array.from(
                    { length: Math.max(1, references) },
                    (_, id) =>
                      ({
                        id,
                        author: "Reviewer",
                        content: [
                          {
                            type: "paragraph",
                            paraId: (0x20000000 + id).toString(16),
                            content: [
                              { type: "run", content: [{ type: "text", text: "Comment" }] },
                            ],
                          },
                        ],
                      }) satisfies Comment,
                  ),
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
            const created = createCanonicalSession(source);
            expect(created.isErr() ? created.error : undefined).toBeUndefined();
            const session = created.unwrap();
            session.setMode(type === "editing" ? { type } : { type, author: "Reviewer" });
            let state = EditorState.create({ doc: session.projection.doc });
            state = state.apply(
              state.tr.setSelection(TextSelection.create(state.doc, 1, state.doc.content.size - 1)),
            );
            const typed = [first, "y", "z"];
            for (const text of typed) {
              const { from, to } = state.selection;
              const prepared = session.prepareReplace(state, {
                from,
                to,
                text,
                semantic: "typing",
              });
              expect(prepared.isErr() ? prepared.error : undefined).toBeUndefined();
              const published = publishCanonicalProjection({
                session,
                state,
                commit: prepared.unwrap(),
              });
              expect(published.isErr() ? published.error : undefined).toBeUndefined();
              state = published.unwrap().state;
              expect(state.selection.empty).toBe(true);
            }
            const { ids: revisionIds } = getTrackedChangeStatsFromDoc(toProseDoc(session.document));
            if (revisionIds.length > 0) {
              const prepared = session.prepareResolve(state, { revisionIds, resolution: "accept" });
              expect(prepared.isErr() ? prepared.error : undefined).toBeUndefined();
              const published = publishCanonicalProjection({
                session,
                state,
                commit: prepared.unwrap(),
              });
              expect(published.isErr() ? published.error : undefined).toBeUndefined();
              state = published.unwrap().state;
            }
            // Comment references retain their anchor but contribute no authored text.
            const projected = toProseDoc(session.document);
            expect(projected.textContent).toBe(typed.join(""));
            const covered: string[] = [];
            projected.descendants((node) => {
              if (
                node.isText &&
                node.marks.some(
                  (mark) => mark.type.name === "comment" && mark.attrs["commentId"] === 0,
                )
              )
                covered.push(node.text ?? "");
            });
            coveredText.set(type, covered.join(""));
          }
          expect(coveredText.get("editing")).toBe("");
          expect(coveredText.get("suggesting")).toBe(first);
        },
      ),
      {
        seed: 1404001,
        numRuns: 12,
        id: "typing after replacement preserves caret order across comment anchors",
      },
    );
  },
  propertyTestTimeout(30_000),
);
