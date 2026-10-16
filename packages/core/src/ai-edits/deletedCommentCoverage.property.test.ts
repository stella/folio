import { expect, test } from "bun:test";
import fc from "fast-check";

import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { createDocx } from "../docx/rezip";
import type { Document, Paragraph, Run } from "../types/document";
import { FolioDocxReviewer } from "./headless";
import { createFolioAITextRangeHandle } from "./snapshot";

// The consumer oracle confused deleted text with a deleted owning block.
// A surviving block retains each comment, collapsing fully deleted ranges to points.
test(
  "pure deletion retains comment coverage or a point in surviving blocks",
  async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.record({
          first: fc.integer({ min: 0, max: 5 }),
          length: fc.integer({ min: 1, max: 6 }),
          cut: fc.constantFrom("whole", "prefix", "suffix", "inside", "outside"),
          location: fc.constantFrom("body", "table"),
          operation: fc.constantFrom("replaceInBlock", "replaceRange"),
          mode: fc.constantFrom("direct", "tracked-changes"),
          count: fc.integer({ min: 1, max: 3 }),
        }),
        async ({ first, length, cut, location, operation, mode, count }) => {
          const source = "abcdefghijkl";
          const last = first + length;
          const cuts = {
            whole: [0, source.length],
            prefix: [0, last],
            suffix: [first, source.length],
            inside: [first, Math.max(first + 1, last - 1)],
            outside: [source.length - 1, source.length],
          } as const;
          const [start, end] = cuts[cut];
          const textRun = (text: string) =>
            ({ type: "run", content: [{ type: "text", text }] }) satisfies Run;
          const ids = Array.from({ length: count }, (_, id) => id);
          const paragraph = {
            type: "paragraph",
            paraId: "12345678",
            content: [
              textRun(source.slice(0, first)),
              ...ids.map((id) => ({ type: "commentRangeStart", id }) as const),
              textRun(source.slice(first, last)),
              ...ids.map((id) => ({ type: "commentRangeEnd", id }) as const),
              ...ids.map((id) => ({ type: "commentReference", id }) as const),
              textRun(source.slice(last)),
            ],
          } satisfies Paragraph;
          const document = {
            package: {
              document: {
                content:
                  location === "body"
                    ? [paragraph]
                    : [
                        {
                          type: "table",
                          rows: [
                            {
                              type: "tableRow",
                              cells: [{ type: "tableCell", content: [paragraph] }],
                            },
                          ],
                        },
                      ],
                comments: ids.map((id) => ({
                  id,
                  author: "Reviewer",
                  content: [{ type: "paragraph", content: [textRun(`Note ${id}`)] }],
                })),
              },
            },
          } satisfies Document;
          const reviewer = await FolioDocxReviewer.fromBuffer(await createDocx(document));
          const block = reviewer.getContent().find(({ text }) => text === source);
          expect(block).toBeDefined();
          if (!block) return;
          const result = reviewer.applyDocumentOperations(
            {
              version: 1,
              mode,
              operations: [
                operation === "replaceInBlock"
                  ? {
                      id: "delete",
                      type: operation,
                      blockId: block.id,
                      find: source.slice(start, end),
                      replace: "",
                    }
                  : {
                      id: "delete",
                      type: operation,
                      range: createFolioAITextRangeHandle({
                        blockId: block.id,
                        text: source,
                        startOffset: start,
                        endOffset: end,
                      }),
                      replace: "",
                    },
              ],
            },
            { undefinedReferences: "refuse" },
          );
          expect(result.applied).toHaveLength(1);
          if (mode === "tracked-changes") {
            expect(reviewer.getComments()).toHaveLength(count);
            reviewer.acceptAll();
          }
          const expected =
            source.slice(first, Math.min(last, start)) + source.slice(Math.max(first, end), last);
          const anchors = (value: FolioDocxReviewer) =>
            value.getComments().map(({ anchoredText }) => anchoredText);
          expect(anchors(reviewer)).toEqual(ids.map(() => expected));
          const reopened = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
          expect(anchors(reopened)).toEqual(anchors(reviewer));
        },
      ),
      { id: "deleted comment coverage follows surviving blocks", numRuns: 64 },
    );
  },
  propertyTestTimeout(60000),
);
