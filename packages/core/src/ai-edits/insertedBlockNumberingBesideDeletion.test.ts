/**
 * A paragraph inserted beside a list item that a pending deletion takes whole
 * takes no numbering from it: once accepted, that list item is gone.
 */
import { describe, expect, test } from "bun:test";

import { paragraphNumberingReference } from "../docx/numberingReference";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { FolioDocxReviewer } from "./headless";

const CLAUSES = ["Numbered clause one", "Numbered clause two"];

const numberedDocument = async (): Promise<ArrayBuffer> => {
  const document = fromMarkdown(
    [
      "# Terms",
      ...CLAUSES.slice(0, 1),
      "Unnumbered body text.",
      ...CLAUSES.slice(1),
    ].join("\n\n"),
  );
  document.package.numbering = {
    abstractNums: [
      {
        abstractNumId: 7,
        multiLevelType: "singleLevel",
        levels: [
          {
            ilvl: 0,
            start: 1,
            numFmt: "decimal",
            lvlText: "(%1)",
            suffix: "space",
            pPr: { indentLeft: 0, indentFirstLine: 0 },
          },
        ],
      },
    ],
    nums: [{ numId: 7, abstractNumId: 7 }],
  };
  for (const block of document.package.document.content)
    if (
      block.type === "paragraph" &&
      block.content.some(
        (node) =>
          node.type === "run" &&
          node.content.some(
            (item) => item.type === "text" && CLAUSES.includes(item.text),
          ),
      )
    )
      block.formatting = {
        ...block.formatting,
        numPr: paragraphNumberingReference({ numId: 7, ilvl: 0 }),
      };
  return createDocx(document);
};

const idOf = (reviewer: FolioDocxReviewer, text: string): string => {
  const block = reviewer.getContent().find((entry) => entry.text === text);
  if (!block) throw new Error(`No block reads ${text}`);
  return block.id;
};

describe.each(["legacy", "canonical"] as const)(
  "on the %s session",
  (session) => {
    for (const mode of ["direct", "tracked-changes"] as const)
      test(`a heading inserted beside a join into a deleted list item takes no numbering from it (${mode})`, async () => {
        const open = async (bytes: ArrayBuffer) =>
          FolioDocxReviewer.fromBuffer(bytes, { session });
        const pending = await open(await numberedDocument());
        const clause = idOf(pending, "Numbered clause two");
        const apply = (
          reviewer: FolioDocxReviewer,
          applied: "direct" | "tracked-changes",
          operations: Parameters<
            FolioDocxReviewer["applyDocumentOperations"]
          >[0]["operations"],
        ) =>
          expect(
            reviewer.applyDocumentOperations(
              { version: 1, mode: applied, operations },
              { undefinedReferences: "refuse" },
            ).skipped,
          ).toEqual([]);
        // The list item goes whole before a table; the body text keeps its
        // words, but its break stays pending, running into the list item.
        apply(pending, "tracked-changes", [
          { id: "delete", type: "deleteBlock", blockId: clause },
          {
            id: "table",
            type: "insertTable",
            blockId: clause,
            rows: [["Term", "Value"]],
          },
        ]);
        const row = pending
          .getChanges()
          .find(({ type }) => type === "rowInserted");
        expect(row && pending.acceptChange(row)).toBe(true);
        apply(pending, "tracked-changes", [
          {
            id: "body",
            type: "deleteBlock",
            blockId: idOf(pending, "Unnumbered body text."),
          },
        ]);
        const words = pending
          .getChanges()
          .find(({ type }) => type === "deletion");
        expect(words && pending.rejectChange(words)).toBe(true);
        const reviewer = await open(await pending.toBuffer());
        apply(reviewer, mode, [
          {
            id: "insert",
            type: "insertAfterBlock",
            blockId: idOf(reviewer, "Unnumbered body text."),
            text: "Inserted heading.",
            styleId: "Heading2",
          },
        ]);
        const inserted = reviewer
          .getContent()
          .find(({ text }) => text === "Inserted heading.");
        expect(inserted?.kind).toBe("heading");
        expect(
          inserted?.kind === "heading" ? inserted.listReference : "missing",
        ).toBeUndefined();
      });
  },
);
