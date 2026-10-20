import { expect, test } from "bun:test";
import fc from "fast-check";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { createDocx } from "../docx/rezip";
import type { Document, Run } from "../types/document";
import { FolioDocxReviewer } from "./headless";

// Whole-paragraph range fixtures missed comments beside unchanged prefix text.
test(
  "aligned comment coverage agrees between direct and accepted tracked replacements",
  async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.integer({ min: 0, max: 22 }),
        fc.integer({ min: 1, max: 24 }),
        fc.constantFrom(
          "Goods agreement goods buyer supplier agreement payment goods.",
          "Goods agreement delivery term.",
          "",
        ),
        async (start, length, replacement) => {
          const source = "Goods agreement delivery.";
          const end = Math.min(source.length, start + length);
          const textRun = (text: string) =>
            ({ type: "run", content: [{ type: "text", text }] }) satisfies Run;
          const ids = Array.from({ length: 1 + (start % 3) }, (_, id) => id);
          const model = {
            package: {
              document: {
                content: [
                  {
                    type: "paragraph",
                    paraId: "12345678",
                    content: [
                      textRun(source.slice(0, start)),
                      ...ids.map((id) => ({ type: "commentRangeStart", id }) as const),
                      textRun(source.slice(start, end)),
                      ...ids.map((id) => ({ type: "commentRangeEnd", id }) as const),
                      ...ids.map((id) => ({ type: "commentReference", id }) as const),
                      textRun(source.slice(end)),
                    ],
                  },
                ],
                comments: ids.map((id) => ({
                  id,
                  author: "Reviewer",
                  content: [
                    {
                      type: "paragraph",
                      paraId: (0x23456789 + id).toString(16),
                      content: [textRun("Note")],
                    },
                  ],
                })),
              },
            },
          } satisfies Document;
          const bytes = await createDocx(model);
          const direct = await FolioDocxReviewer.fromBuffer(bytes);
          const tracked = await FolioDocxReviewer.fromBuffer(bytes);
          for (const [reviewer, mode] of [
            [direct, "direct"],
            [tracked, "tracked-changes"],
          ] as const) {
            const target = reviewer.getContent().at(0);
            if (!target) throw new Error("Expected paragraph");
            const receipt = reviewer.applyDocumentOperations(
              {
                version: 1,
                mode,
                operations: [
                  { id: "replace", type: "replaceBlock", blockId: target.id, text: replacement },
                ],
              },
              { undefinedReferences: "refuse" },
            );
            expect(receipt.applied).toHaveLength(1);
          }
          const pending = await tracked.toBuffer();
          tracked.acceptAll();
          const comments = (reviewer: FolioDocxReviewer) =>
            reviewer
              .getComments()
              .map(({ id, text, anchoredText }) => ({ id, text, anchoredText }));
          expect(comments(tracked)).toEqual(comments(direct));
          expect(comments(direct)).toHaveLength(replacement === "" ? 0 : ids.length);
          const reopenedDirect = await FolioDocxReviewer.fromBuffer(await direct.toBuffer());
          const reopenedTracked = await FolioDocxReviewer.fromBuffer(await tracked.toBuffer());
          expect(comments(reopenedTracked)).toEqual(comments(reopenedDirect));
          expect(comments(reopenedDirect)).toEqual(comments(direct));
          const individually = await FolioDocxReviewer.fromBuffer(pending);
          for (const change of individually.getChanges()) individually.acceptChange(change);
          expect(comments(individually)).toEqual(comments(direct));
          const rejected = await FolioDocxReviewer.fromBuffer(pending);
          rejected.rejectAll();
          expect(rejected.getComments().at(0)?.anchoredText).toBe(source.slice(start, end));
          if (replacement === "") expect(direct.getComments().at(0)?.anchoredText ?? "").toBe("");
        },
      ),
      {
        id: "partial comment spans retain aligned coverage and point anchors in both edit modes",
        seed: 1683001,
        numRuns: 16,
      },
    );
  },
  propertyTestTimeout(60000),
);
