import { expect, test } from "bun:test";
import fc from "fast-check";
import { AllSelection, EditorState } from "prosemirror-state";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { canonicalSelectionRange } from "../prosemirror/canonicalSelectionRange";
import type { Paragraph } from "../types/document";
import { createCanonicalSession } from "./canonicalSession";

test(
  "every canonical input owner resolves document edges as content gaps",
  () => {
    assertProperty(
      fc.property(
        fc.array(
          fc
            .array(fc.constantFrom("a", "😀", "文"), { maxLength: 8 })
            .map((characters) => characters.join("")),
          { minLength: 1, maxLength: 4 },
        ),
        (texts) => {
          const session = createCanonicalSession({
            package: {
              document: {
                content: texts.map(
                  (text, index) =>
                    ({
                      type: "paragraph",
                      paraId: (0x10000000 + index).toString(16).toUpperCase(),
                      content: [
                        { type: "run", content: [{ type: "text", text }] },
                      ],
                    }) satisfies Paragraph,
                ),
              },
            },
          }).unwrap();
          const doc = session.projection.doc;
          const state = EditorState.create({
            doc,
            selection: new AllSelection(doc),
          });
          const range = canonicalSelectionRange(state);
          for (const [edge, content] of [
            [0, range.from],
            [doc.content.size, range.to],
          ] as const) {
            expect(session.projection.addressAt(edge).unwrap()).toEqual(
              session.projection.addressAt(content).unwrap(),
            );
            expect(session.projection.inputAddressAt(edge).unwrap()).toEqual(
              session.projection.inputAddressAt(content).unwrap(),
            );
          }
          expect(session.projection.addressAt(-1).isErr()).toBe(true);
          expect(
            session.projection.addressAt(doc.content.size + 1).isErr(),
          ).toBe(true);
          let boundary = 0;
          for (let index = 0; index < doc.childCount - 1; index++) {
            boundary += doc.child(index).nodeSize;
            expect(session.projection.addressAt(boundary).isErr()).toBe(true);
          }
        },
      ),
      { numRuns: 30, id: "canonical document edges denote content gaps" },
    );
  },
  propertyTestTimeout(30_000),
);
