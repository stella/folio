/** Numbering edits preserve a heading's outline and inherited run defaults. */

import { expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { FolioDocxReviewer } from "./headless";

const TITLE = "Heading text";
const INSERTED = "Inserted heading";
const STYLE_ID = "NumberedHeading";

const generatedCase = fc.record({
  styleLevel: fc.integer({ min: 0, max: 8 }),
  directLevel: fc.option(fc.integer({ min: 0, max: 8 }), { nil: undefined }),
  bold: fc.boolean(),
  italic: fc.boolean(),
  styleNumbered: fc.boolean(),
  mode: fc.constantFrom("direct", "tracked-changes", "suggested"),
  levels: fc.array(fc.integer({ min: 0, max: 8 }), { minLength: 1, maxLength: 3 }),
});

const headingState = (reviewer: FolioDocxReviewer, text: string) => {
  const block = reviewer.getContent().find((candidate) => candidate.text === text);
  if (!block) throw new Error("heading fixture must expose its text");
  return {
    kind: block.kind,
    headingLevel: block.headingLevel,
    runs: block.previewRuns?.map((run) => ({
      text: run.text,
      bold: run.bold === true,
      italic: run.italic === true,
    })),
  };
};

const expectReopenedHeading = async (reviewer: FolioDocxReviewer, text: string) => {
  const before = headingState(reviewer, text);
  const reopened = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
  expect(headingState(reopened, text)).toEqual(before);
};

test(
  "heading levels and run defaults survive numbering edits and save",
  async () => {
    await fc.assert(
      fc.asyncProperty(generatedCase, async (entry) => {
        const model = fromMarkdown(TITLE);
        model.package.styles?.styles.push({
          styleId: STYLE_ID,
          type: "paragraph",
          name: STYLE_ID,
          pPr: {
            outlineLevel: { kind: "heading", level: entry.styleLevel },
            ...(entry.styleNumbered && {
              numPr: { kind: "reference", numId: 7, ilvl: 0 },
            }),
          },
          rPr: { bold: entry.bold, italic: entry.italic },
        });
        model.package.numbering = {
          abstractNums: [
            {
              abstractNumId: 7,
              multiLevelType: "multilevel",
              levels: Array.from({ length: 9 }, (_, ilvl) => ({
                ilvl,
                start: 1,
                numFmt: "decimal",
                lvlText: `%${ilvl + 1}.`,
              })),
            },
          ],
          nums: [{ numId: 7, abstractNumId: 7 }],
        };
        const paragraph = model.package.document.content.at(0);
        if (paragraph?.type !== "paragraph") throw new Error("fixture must begin with a paragraph");
        paragraph.formatting = {
          styleId: STYLE_ID,
          ...(entry.directLevel !== undefined && {
            outlineLevel: { kind: "heading", level: entry.directLevel },
          }),
        };
        const reviewer = await FolioDocxReviewer.fromBuffer(await createDocx(model), {
          author: "Editor",
        });
        const block = reviewer.getContent().find(({ text }) => text === TITLE);
        if (!block) throw new Error("fixture must expose its heading");
        const expected = headingState(reviewer, TITLE);
        expect(expected.headingLevel).toBe((entry.directLevel ?? entry.styleLevel) + 1);
        expect(expected.runs).toEqual([{ text: TITLE, bold: entry.bold, italic: entry.italic }]);

        for (const level of entry.levels) {
          const result = reviewer.applyDocumentOperations({
            version: 1,
            mode: entry.mode,
            operations: [
              {
                id: "level",
                type: "setBlockParagraphProperties",
                blockId: block.id,
                properties: { numbering: { numId: 7, level } },
              },
            ],
          });
          expect(result.skipped).toEqual([]);
          expect(headingState(reviewer, TITLE)).toEqual(expected);
          // Pending suggestions are excluded from the package, but numbering
          // must leave the heading's level and text formatting equal either way.
          await expectReopenedHeading(reviewer, TITLE);
          if (entry.mode === "suggested") reviewer.acceptSuggestion("level");
          reviewer.acceptAll();
          expect(headingState(reviewer, TITLE)).toEqual(expected);
        }
        const inserted = reviewer.applyDocumentOperations({
          version: 1,
          mode: entry.mode,
          operations: [
            {
              id: "insert",
              type: "insertAfterBlock",
              blockId: block.id,
              text: INSERTED,
              styleId: STYLE_ID,
            },
          ],
        });
        expect(inserted.skipped).toEqual([]);
        if (entry.mode === "suggested") expect(reviewer.acceptSuggestion("insert")).toBe(true);
        reviewer.acceptAll();
        expect(headingState(reviewer, TITLE)).toEqual(expected);
        expect(headingState(reviewer, INSERTED)).toEqual({
          ...expected,
          runs: [{ text: INSERTED, bold: entry.bold, italic: entry.italic }],
        });
        await expectReopenedHeading(reviewer, TITLE);
        await expectReopenedHeading(reviewer, INSERTED);
      }),
      propertyConfig({ numRuns: 30 }),
    );
  },
  propertyTestTimeout(30_000),
);
