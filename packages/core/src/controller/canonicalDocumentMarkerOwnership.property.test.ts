import { expect, test } from "bun:test";
import fc from "fast-check";
import { AllSelection, EditorState } from "prosemirror-state";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { getTrackedChangeStatsFromDoc } from "../ai-edits/read";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Document, Paragraph } from "../types/document";
import {
  createCanonicalSession,
  publishCanonicalProjection,
} from "./canonicalSession";
import { prepareCanonicalAutoformat } from "./canonicalStructure";

// Marker fixtures started in one paragraph and never retired the addressed mark.
test(
  "document replacement markers format the surviving paragraph mark in both modes",
  () => {
    assertProperty(
      fc.property(
        fc.integer({ min: 1, max: 4 }),
        fc.constantFrom("-", "*", "1.", "#"),
        (count, marker) => {
          const source = {
            package: {
              document: {
                content: Array.from(
                  { length: count },
                  (_, index) =>
                    ({
                      type: "paragraph",
                      paraId: (0x12345678 + index).toString(16),
                      content: [
                        {
                          type: "run",
                          content: [
                            { type: "text", text: `Paragraph ${index}` },
                          ],
                        },
                      ],
                    }) satisfies Paragraph,
                ),
              },
              styles: {
                styles: [
                  { styleId: "Heading1", type: "paragraph", name: "Heading 1" },
                ],
              },
            },
          } satisfies Document;
          const outcomes = [];
          for (const mode of ["editing", "suggesting"] as const) {
            const session = createCanonicalSession(source).unwrap();
            session.setMode(
              mode === "editing"
                ? { type: mode }
                : { type: mode, author: "Reviewer" },
            );
            let state = EditorState.create({ doc: session.projection.doc });
            state = state.apply(
              state.tr.setSelection(new AllSelection(state.doc)),
            );
            for (const text of [...marker, " "]) {
              const input = {
                from: state.selection.from,
                to: state.selection.to,
                text,
              };
              const commit = (
                prepareCanonicalAutoformat(session, state, input) ??
                session.prepareReplace(state, { ...input, semantic: "typing" })
              ).unwrap();
              state = publishCanonicalProjection({
                session,
                state,
                commit,
              }).unwrap().state;
            }
            const { ids } = getTrackedChangeStatsFromDoc(
              toProseDoc(session.document),
            );
            if (ids.length > 0) {
              const commit = session
                .prepareResolve(state, {
                  revisionIds: ids,
                  resolution: "accept",
                })
                .unwrap();
              state = publishCanonicalProjection({
                session,
                state,
                commit,
              }).unwrap().state;
            }
            expect(session.document.package.document.content).toHaveLength(1);
            const paragraph = session.document.package.document.content.at(0);
            if (paragraph?.type !== "paragraph")
              throw new Error("Expected paragraph");
            expect(toProseDoc(session.document).textContent).toBe("");
            outcomes.push({
              numPr: paragraph.formatting?.numPr,
              styleId: paragraph.formatting?.styleId,
            });
            if (marker === "#")
              expect(paragraph.formatting?.styleId).toBe("Heading1");
            else expect(paragraph.formatting?.numPr?.kind).toBe("reference");
          }
          expect(outcomes.at(1)).toEqual(outcomes.at(0));
        },
      ),
      {
        id: "document marker rules retain ending paragraph property ownership",
        seed: 1404003,
        numRuns: 16,
      },
    );
  },
  propertyTestTimeout(30_000),
);
