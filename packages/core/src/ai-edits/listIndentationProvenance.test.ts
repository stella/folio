import { expect, test } from "bun:test";
import { expectParagraphBlock } from "../../../../test/paragraphBlock";

import { createDocx } from "../docx/rezip";
import { getParagraphText } from "../docx/paragraphParser";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { paragraphNumberingReference } from "../docx/numberingReference";
import { FolioDocxReviewer } from "./headless";
import { isFolioAIContentBlock } from "./snapshot";

const CASES = (["direct", "tracked-changes"] as const).flatMap((mode) =>
  (["set-properties", "insert", "remove-properties", "insert-removed"] as const).flatMap(
    (operation) =>
      (["inherited", "zero", "explicit"] as const).flatMap((source) =>
        (["preserve", "clear", "zero"] as const).flatMap((indentation) =>
          (operation === "remove-properties" || operation === "insert-removed"
            ? [false, true]
            : [true]
          ).map((withStyle) => ({ mode, operation, source, indentation, withStyle })),
        ),
      ),
  ),
);

// Separate style and numbering examples missed their combined write order:
// the later style patch replaced the numbering-derived provenance.
test.each(CASES)(
  "$operation $source/$indentation indentation survives numbering (style=$withStyle) in $mode",
  async ({ mode, operation, source, indentation, withStyle }) => {
    const document = fromMarkdown("1. Anchor");
    const paragraph = document.package.document.content.at(0);
    const definition = document.package.numbering?.abstractNums.at(0);
    const template = definition?.levels.at(0);
    const numPr = paragraph?.type === "paragraph" ? paragraph.formatting?.numPr : undefined;
    if (paragraph?.type !== "paragraph" || !definition || !template || numPr?.kind !== "reference")
      throw new TypeError("Combined list fixture lacks numbering.");
    const styles = document.package.styles;
    if (!styles) throw new TypeError("Combined list fixture lacks styles.");
    styles.styles.push({
      type: "paragraph",
      styleId: "IndentBaseline",
      name: "Indent baseline",
      pPr: { indentLeft: 100, indentFirstLine: 20, hangingIndent: false },
    });
    definition.levels = [0, 1].map((ilvl) =>
      Object.assign({}, template, {
        ilvl,
        lvlText: `%${String(ilvl + 1)}.`,
        pPr: { indentLeft: 720 * (ilvl + 1), indentFirstLine: -360, hangingIndent: true },
      }),
    );
    const sourceDirect =
      source === "inherited"
        ? undefined
        : {
            indentLeft: source === "zero" ? 0 : 1440,
            indentFirstLine: 0,
            hangingIndent: false,
          };
    paragraph.paraId = "12345678";
    paragraph.formatting = {
      numPr: paragraphNumberingReference({ numId: numPr.numId, ilvl: 0 }),
      ...sourceDirect,
    };
    const reviewer = await FolioDocxReviewer.fromBuffer(await createDocx(document));
    const anchor = reviewer.getContent().find(({ text }) => text === "Anchor");
    if (!anchor) throw new TypeError("Combined list fixture lost its anchor.");
    const explicitIndentation = (() => {
      switch (indentation) {
        case "preserve":
          return undefined;
        case "clear":
          return null;
        case "zero":
          return { indentLeft: 0, indentFirstLine: 0, hangingIndent: false };
      }
    })();
    const expectedDirect =
      indentation === "preserve" ? sourceDirect : (explicitIndentation ?? undefined);
    const removesNumbering = operation === "remove-properties" || operation === "insert-removed";
    const inserts = operation === "insert" || operation === "insert-removed";
    const properties = {
      ...(withStyle && { styleId: "IndentBaseline" }),
      numbering: removesNumbering
        ? { kind: "none" as const }
        : { kind: "reference" as const, numId: numPr.numId, ilvl: 1 },
      ...(explicitIndentation !== undefined && { indentation: explicitIndentation }),
    };
    const result = reviewer.applyDocumentOperations(
      {
        version: 1,
        mode,
        operations: [
          !inserts
            ? {
                id: "change",
                type: "setBlockParagraphProperties",
                blockId: anchor.id,
                properties,
              }
            : {
                id: "change",
                type: "insertAfterBlock",
                blockId: anchor.id,
                text: "Inserted",
                ...properties,
              },
        ],
      },
      { undefinedReferences: "refuse" },
    );
    expect(result.status).toBe("committed");
    expect(result.issues).toEqual([]);
    const text = inserts ? "Inserted" : "Anchor";
    const assertContent = (current: FolioDocxReviewer, review: "pending" | "settled") => {
      const observedBlock = current
        .getContent()
        .filter(isFolioAIContentBlock)
        .find((entry) => entry.text === text);
      if (!observedBlock) throw new TypeError("Combined list operation lost its target.");
      const block = expectParagraphBlock(observedBlock);
      expect(block.styleId).toBe(withStyle ? "IndentBaseline" : anchor.styleId);
      if (removesNumbering && review === "settled") expect(block.displayLabel).toBeUndefined();
      const removedLabelReference =
        block.displayLabel === undefined ? undefined : { numId: numPr.numId, level: 0 };
      expect(block.listReference).toEqual(
        removesNumbering ? removedLabelReference : { numId: numPr.numId, level: 1 },
      );
      expect(block.directIndentation).toEqual(expectedDirect);
    };
    const review = mode === "tracked-changes" ? "pending" : "settled";
    assertContent(reviewer, review);
    assertContent(await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer()), review);
    if (mode === "tracked-changes") {
      expect(reviewer.acceptAll()).toBeGreaterThan(0);
      assertContent(reviewer, "settled");
      assertContent(await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer()), "settled");
    }
  },
);

test.each([2, 3, 4, 5, 6, 7, 8])(
  "undefined numbering level %i omits a redundant false hanging indent",
  async (ilvl) => {
    const document = fromMarkdown(
      [
        "# Agreement",
        "## Scope",
        "### Definitions",
        "The Supplier delivers the goods.",
        "The Buyer pays on delivery.",
        "## Payment",
        "Payment is due in ten days.",
        "Closing paragraph.",
        "1. Anchor",
      ].join("\n\n"),
    );
    const definition = document.package.numbering?.abstractNums.at(0);
    const level = definition?.levels.at(0);
    const numberedParagraph = document.package.document.content.find(
      (block) => block.type === "paragraph" && getParagraphText(block) === "Anchor",
    );
    const numPr =
      numberedParagraph?.type === "paragraph" ? numberedParagraph.formatting?.numPr : undefined;
    const numbering = document.package.numbering;
    const heading2 = document.package.styles?.styles.find(
      (style) => style.type === "paragraph" && style.styleId === "Heading2",
    );
    const heading3 = document.package.styles?.styles.find(
      (style) => style.type === "paragraph" && style.styleId === "Heading3",
    );
    if (
      !definition ||
      !level ||
      numPr?.kind !== "reference" ||
      !numbering ||
      heading2?.type !== "paragraph" ||
      heading3?.type !== "paragraph"
    ) {
      throw new TypeError("Numbered fixture lacks its numbering definition.");
    }
    definition.levels = [
      { ...level, ilvl: 0, pPr: { indentLeft: 0, indentFirstLine: 0 } },
      { ...level, ilvl: 1, pPr: { indentLeft: 0, indentFirstLine: 0 } },
    ];
    heading2.pPr = {
      ...heading2.pPr,
      numPr: paragraphNumberingReference({ numId: numPr.numId, ilvl: 0 }),
    };
    heading3.pPr = {
      ...heading3.pPr,
      numPr: paragraphNumberingReference({ numId: numPr.numId, ilvl: 1 }),
    };

    const reviewer = await FolioDocxReviewer.fromBuffer(await createDocx(document));
    const closing = reviewer.getContent().find(({ text }) => text === "Closing paragraph.");
    const supplier = reviewer
      .getContent()
      .find(({ text }) => text === "The Supplier delivers the goods.");
    if (!closing || !supplier) throw new TypeError("Numbered fixture lost an insertion anchor.");
    expect(
      reviewer
        .applyDocumentOperations(
          {
            version: 1,
            mode: "tracked-changes",
            operations: [
              {
                id: "list-insertion",
                type: "insertAfterBlock",
                blockId: supplier.id,
                text: "Warranty period period.",
                numbering: paragraphNumberingReference({ numId: numPr.numId, ilvl: 0 }),
              },
              {
                id: "styled-insertion",
                type: "insertAfterBlock",
                blockId: closing.id,
                text: "Supplier agreement.",
                styleId: "Heading2",
              },
            ],
          },
          { undefinedReferences: "refuse" },
        )
        .applied.map(({ id }) => id),
    ).toEqual(["styled-insertion", "list-insertion"]);
    const styledInsertion = reviewer
      .getContent()
      .find(({ text }) => text === "Supplier agreement.");
    if (!styledInsertion) throw new TypeError("Tracked style insertion was lost.");
    const result = reviewer.applyDocumentOperations(
      {
        version: 1,
        mode: "tracked-changes",
        operations: [
          {
            id: "undefined-level",
            type: "insertAfterBlock",
            blockId: styledInsertion.id,
            text: "x",
            numbering: paragraphNumberingReference({ numId: numPr.numId, ilvl }),
          },
        ],
      },
      { undefinedReferences: "refuse" },
    );
    expect(result.applied.map(({ id }) => id)).toEqual(["undefined-level"]);
    const expectedReviewer = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
    const expectedDocument = expectedReviewer.toResolvedDocument("accept");
    const expectedInserted = expectedDocument.package.document.content.find(
      (block) => block.type === "paragraph" && getParagraphText(block) === "x",
    );
    if (expectedInserted?.type !== "paragraph") {
      throw new TypeError("Accepted save lost the undefined-level paragraph.");
    }

    const actualDocument = reviewer.toResolvedDocument("accept");
    const actualInserted = actualDocument.package.document.content.find(
      (block) => block.type === "paragraph" && getParagraphText(block) === "x",
    );
    if (actualInserted?.type !== "paragraph") {
      throw new TypeError("Accepted live state lost the undefined-level paragraph.");
    }
    expect(actualInserted.formatting?.hangingIndent).toBeUndefined();
    expect(actualInserted.formatting).toEqual(expectedInserted.formatting);
  },
);
