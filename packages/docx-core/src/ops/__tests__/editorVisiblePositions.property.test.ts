import { expect, test } from "bun:test";
import fc from "fast-check";
import { assertProperty, propertyTestTimeout } from "../../../../../test/property-testing";
import type { Document, Paragraph } from "../../model/document";
import {
  editorParagraphGroups,
  editorPositionAtPhysicalPosition,
  physicalPositionAtEditorOffset,
} from "../editorIntent";
import { OP_STORIES } from "../types";

// Physical offsets counted deleted characters; marker rules need visible offsets.
test(
  "logical paragraph gaps round-trip through retained deleted source text",
  () => {
    assertProperty(
      fc.property(
        fc.array(
          fc.record({
            kept: fc
              .array(fc.constantFrom("a", "😀", "文"), { maxLength: 4 })
              .map((text) => text.join("")),
            removed: fc
              .array(fc.constantFrom("b", "👍🏽"), { maxLength: 4 })
              .map((text) => text.join("")),
          }),
          { minLength: 1, maxLength: 4 },
        ),
        (parts) => {
          const document = {
            package: {
              document: {
                content: parts.map(
                  ({ kept, removed }, index) =>
                    ({
                      type: "paragraph",
                      paraId: (0x12345678 + index).toString(16),
                      ...(index + 1 < parts.length
                        ? {
                            pPrMark: { kind: "del", info: { id: index + 100, author: "Reviewer" } },
                          }
                        : {}),
                      content: [
                        { type: "run", content: [{ type: "text", text: kept }] },
                        {
                          type: "deletion",
                          info: { id: index + 1, author: "Reviewer" },
                          content: [{ type: "run", content: [{ type: "text", text: removed }] }],
                        },
                      ],
                    }) satisfies Paragraph,
                ),
              },
            },
          } satisfies Document;
          const group = editorParagraphGroups(document, OP_STORIES.MAIN).at(0);
          if (!group) throw new Error("Expected logical paragraph");
          expect(group.text).toBe(parts.map(({ kept }) => kept).join(""));
          for (let offset = 0; offset <= group.text.length; offset++) {
            const logical = { story: OP_STORIES.MAIN, blockId: group.blockId, offset };
            const physical = physicalPositionAtEditorOffset(document, logical);
            expect(editorPositionAtPhysicalPosition(document, physical)).toEqual(logical);
          }
        },
      ),
      {
        id: "logical gaps map through deleted text and retired paragraph marks",
        seed: 1404004,
        numRuns: 24,
      },
    );
  },
  propertyTestTimeout(30_000),
);
