import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { createDocx } from "../docx/rezip";
import type { Document, ParagraphContent, RunContent } from "../types/document";
import {
  openReviewer,
  paragraphsDocx,
  reviewComment,
  textRun,
} from "../__tests__/operationBatchDocuments";

const COMMENT_ID = 7;
const PNG_DATA_URL =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

const image: RunContent = {
  type: "drawing",
  image: {
    type: "image",
    src: PNG_DATA_URL,
    size: { width: 914_400, height: 914_400 },
    wrap: { type: "inline" },
  },
};

const atomContent = {
  image: { type: "run", content: [image] },
  field: {
    type: "simpleField",
    fieldType: "REF",
    instruction: " REF _Ref1 \\h ",
    content: [],
  },
  tab: { type: "run", content: [{ type: "tab" }] },
  noteMarker: { type: "run", content: [{ type: "noteMarker", kind: "footnote" }] },
} as const satisfies Record<string, ParagraphContent>;

const commentedParagraph = (
  covered: readonly ParagraphContent[],
  { reference = true }: { reference?: boolean } = {},
): ParagraphContent[] => [
  { type: "commentRangeStart", id: COMMENT_ID },
  ...covered,
  { type: "commentRangeEnd", id: COMMENT_ID },
  ...(reference ? [{ type: "commentReference" as const, id: COMMENT_ID }] : []),
];

const deleteGone = async (
  covered: readonly ParagraphContent[],
  mode: "direct" | "tracked-changes",
  resolution: "all" | "single" = "all",
  deleted = ["gone"],
) => {
  const reviewer = await openReviewer(
    await paragraphsDocx([commentedParagraph(covered)], [reviewComment(COMMENT_ID)]),
  );
  const blockId = reviewer.snapshot().blocks.at(0)?.id;
  if (!blockId) throw new Error("Expected a paragraph block");
  const result = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode,
      operations: deleted.map((find, index) => ({
        id: `delete-${String(index)}`,
        type: "replaceInBlock" as const,
        blockId,
        find,
        replace: "",
      })),
    },
    { undefinedReferences: "refuse" },
  );
  expect(result.skipped).toEqual([]);
  if (mode === "tracked-changes") {
    if (resolution === "all") reviewer.acceptAll();
    else {
      const change = reviewer.getChanges().at(0);
      expect(change).toBeDefined();
      expect(reviewer.acceptChange(change!)).toBe(true);
    }
  }
  const reopened = await openReviewer(await reviewer.toBuffer());
  return reopened.getComments().some(({ id }) => id === COMMENT_ID);
};

const rejectInserted = async (survivors: readonly ParagraphContent[]) => {
  const reviewer = await openReviewer(
    await paragraphsDocx(
      [
        commentedParagraph([
          {
            type: "insertion",
            info: { id: 19, author: "Reviewer" },
            content: [textRun("gone")],
          },
          ...survivors,
        ]),
      ],
      [reviewComment(COMMENT_ID)],
    ),
  );
  expect(reviewer.rejectAll()).toBeGreaterThan(0);
  return reviewer.getComments().some(({ id }) => id === COMMENT_ID);
};

describe("accepted deletion comment anchors", () => {
  test("content in an accepted deleted row does not keep an otherwise exhausted range", async () => {
    const info = { id: 19, author: "Reviewer" };
    const document = {
      package: {
        document: {
          comments: [reviewComment(COMMENT_ID)],
          content: [
            {
              type: "paragraph",
              paraId: "10000000",
              content: [
                { type: "commentRangeStart", id: COMMENT_ID },
                { type: "deletion", info, content: [textRun("gone")] },
              ],
            },
            {
              type: "table",
              rows: [
                {
                  type: "tableRow",
                  structuralChange: { type: "tableRowDeletion", info: { ...info, id: 20 } },
                  cells: [
                    {
                      type: "tableCell",
                      content: [
                        {
                          type: "paragraph",
                          paraId: "10000001",
                          content: [textRun("removed with row")],
                        },
                      ],
                    },
                  ],
                },
              ],
            },
            {
              type: "paragraph",
              paraId: "10000002",
              content: [
                { type: "commentRangeEnd", id: COMMENT_ID },
                { type: "commentReference", id: COMMENT_ID },
              ],
            },
          ],
        },
      },
    } satisfies Document;
    const reviewer = await openReviewer(await createDocx(document));

    expect(reviewer.acceptAll()).toBeGreaterThan(0);
    expect(reviewer.getComments().some(({ id }) => id === COMMENT_ID)).toBe(false);
  });

  test.each(["all", "single"] as const)(
    "a retained inline image keeps the thread after accepting %s revisions",
    async (resolution) => {
      expect(
        await deleteGone([textRun("gone"), atomContent.image], "tracked-changes", resolution),
      ).toBe(true);
    },
  );

  test.each(["field", "tab"] as const)("a retained inline %s keeps the thread", async (kind) => {
    expect(await deleteGone([textRun("gone"), atomContent[kind]], "tracked-changes")).toBe(true);
  });

  test("a surviving block loses a range when its last covered text is deleted", async () => {
    expect(await deleteGone([textRun("gone")], "tracked-changes")).toBe(false);
  });

  test("direct deletion removes the same exhausted range", async () => {
    expect(await deleteGone([textRun("gone")], "direct")).toBe(false);
  });

  test(
    "a range survives exactly when covered non-carrier content survives",
    async () => {
      const property = fc.asyncProperty(
        fc.record({
          deletedTextCount: fc.integer({ min: 1, max: 4 }),
          survivors: fc.array(fc.constantFrom("text", "image", "field", "tab", "noteMarker"), {
            maxLength: 6,
          }),
        }),
        async ({ deletedTextCount, survivors }) => {
          const covered: ParagraphContent[] = [];
          const deleted: string[] = [];
          for (let index = 0; index < deletedTextCount; index += 1) {
            const value = `deleted${String(index)}`;
            covered.push(textRun(value));
            deleted.push(value);
          }
          for (const [index, kind] of survivors.entries()) {
            covered.push(kind === "text" ? textRun(`survivor${String(index)}`) : atomContent[kind]);
          }
          const expected = survivors.length > 0;
          expect(await deleteGone(covered, "direct", "all", deleted)).toBe(expected);
          expect(await deleteGone(covered, "tracked-changes", "all", deleted)).toBe(expected);
          expect(
            await rejectInserted(
              survivors.map((kind, index) =>
                kind === "text" ? textRun(`survivor${String(index)}`) : atomContent[kind],
              ),
            ),
          ).toBe(expected);
        },
      );
      await fc.assert(property, propertyConfig({ numRuns: 40, seed: 1958 }));
    },
    propertyTestTimeout(60_000),
  );
});

describe("rejected insertion comment anchors", () => {
  test("rejecting inserted text retires its thread and source anchors", async () => {
    const reviewer = await openReviewer(await paragraphsDocx([[textRun("Original")]]));
    const blockId = reviewer.snapshot().blocks.at(0)?.id;
    if (!blockId) throw new Error("Expected a paragraph block");
    const inserted = reviewer.applyOperations(
      [{ id: "insert", type: "insertAfterBlock", blockId, text: "Temporary" }],
      { undefinedReferences: "refuse", mode: "tracked-changes" },
    );
    expect(inserted.skipped).toEqual([]);
    const temporary = reviewer.snapshot().blocks.find(({ text }) => text === "Temporary");
    expect(temporary).toBeDefined();
    const commented = reviewer.applyOperations(
      [
        {
          id: "comment",
          type: "commentOnBlock",
          blockId: temporary!.id,
          comment: { text: "Remove with insertion" },
        },
      ],
      { undefinedReferences: "refuse", mode: "direct" },
    );
    expect(commented.skipped).toEqual([]);

    expect(reviewer.rejectAll()).toBeGreaterThan(0);
    expect(reviewer.getComments()).toEqual([]);
    expect((await openReviewer(await reviewer.toBuffer())).getComments()).toEqual([]);
  });

  test("rejecting inserted text keeps a range over surviving original text", async () => {
    const reviewer = await openReviewer(await paragraphsDocx([[textRun("Original")]]));
    const blockId = reviewer.snapshot().blocks.at(0)?.id;
    if (!blockId) throw new Error("Expected a paragraph block");
    const inserted = reviewer.applyOperations(
      [
        {
          id: "insert",
          type: "replaceInBlock",
          blockId,
          find: "Original",
          replace: "Temporary Original",
        },
      ],
      { undefinedReferences: "refuse", mode: "tracked-changes" },
    );
    expect(inserted.skipped).toEqual([]);
    const commented = reviewer.applyOperations(
      [
        {
          id: "comment",
          type: "commentOnBlock",
          blockId,
          comment: { text: "Keep on original" },
        },
      ],
      { undefinedReferences: "refuse", mode: "direct" },
    );
    expect(commented.skipped).toEqual([]);

    expect(reviewer.rejectAll()).toBeGreaterThan(0);
    expect(
      reviewer.getComments().map(({ text, anchoredText }) => ({ text, anchoredText })),
    ).toEqual([{ text: "Keep on original", anchoredText: "Original" }]);
  });
});
