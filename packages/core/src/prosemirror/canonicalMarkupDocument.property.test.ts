import { expect, test } from "bun:test";
import fc from "fast-check";
import { EditorState } from "prosemirror-state";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { resolveStateStory } from "../__tests__/reference-editor/markupViewProjection";
import { canonicalMarkupDocument } from "./canonicalMarkupDocument";
import { markupViewPositionMap } from "./markupViewPositionMap";
import { schema } from "./schema";

const runArbitrary = fc.record({
  text: fc.constantFrom("a", "aa", "b", " a ", "🙂"),
  revision: fc.constantFrom("plain", "insertion", "deletion"),
});

test(
  "canonical markup resolution reads the same text as the reference",
  () => {
    assertProperty(
      fc.property(
        fc.array(fc.array(runArbitrary, { minLength: 1, maxLength: 6 }), {
          minLength: 1,
          maxLength: 6,
        }),
        (paragraphs) => {
          let id = 1;
          const doc = schema.node(
            "doc",
            null,
            paragraphs.map((runs) =>
              schema.node(
                "paragraph",
                null,
                runs.map(({ text, revision }) =>
                  schema.text(
                    text,
                    revision === "plain"
                      ? []
                      : [
                          schema.mark(revision, {
                            revisionId: id++,
                            author: "Reviewer",
                          }),
                        ],
                  ),
                ),
              ),
            ),
          );
          const state = EditorState.create({ doc });
          for (const decision of ["accept", "reject"] as const) {
            const reference = resolveStateStory(state, decision);
            const candidate = canonicalMarkupDocument({ state, decision });
            expect(
              candidate.doc.textBetween(0, candidate.doc.content.size, "\n"),
            ).toBe(
              reference.resolved.textBetween(
                0,
                reference.resolved.content.size,
                "\n",
              ),
            );
            expect(candidate.partial).toBe(reference.failed);
            const positions =
              candidate.positionMap ??
              markupViewPositionMap({
                before: doc,
                after: candidate.doc,
                decision,
              });
            for (let position = 0; position <= doc.content.size; position++) {
              for (const assoc of [-1, 1] as const) {
                expect(positions.mapResult(position, assoc)).toEqual(
                  reference.positionMap.mapResult(position, assoc),
                );
              }
            }
          }
        },
      ),
      {
        id: "canonical markup resolution reads the same text as the reference",
        numRuns: 80,
      },
    );
  },
  propertyTestTimeout(10_000),
);
