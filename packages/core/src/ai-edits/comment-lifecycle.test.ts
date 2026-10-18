import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig } from "../../../../test/property-testing";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import type { ParagraphContent, RunContent } from "../types/content";
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

describe("accepted deletion comment anchors", () => {
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

  test("a surviving block keeps a point when its last text is deleted", async () => {
    expect(await deleteGone([textRun("gone")], "tracked-changes")).toBe(true);
  });

  test("surviving content and point fallbacks retain the thread", async () => {
    const property = fc.asyncProperty(
      fc.record({
        deletedTextCount: fc.integer({ min: 1, max: 4 }),
        survivors: fc.array(fc.constantFrom("text", "image", "field", "tab"), {
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
        expect(await deleteGone(covered, "tracked-changes", "all", deleted)).toBe(true);
      },
    );
    await fc.assert(property, propertyConfig({ numRuns: 40, seed: 1958 }));
  });
});
