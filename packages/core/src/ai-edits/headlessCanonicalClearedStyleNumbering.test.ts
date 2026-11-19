/**
 * A paragraph numbered through its style keeps no numbering once an edit
 * clears or changes the style: the canonical model, its accepted view and the
 * reopened save agree, for the paragraph edited and for one inserted with a
 * style of its own beside it.
 */
import { expect, test } from "bun:test";

import { paragraphNumberingFromSlots } from "@stll/docx-core/model";

import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import type { Document, Paragraph } from "../types/document";
import { FolioDocxReviewer } from "./headless";

const source = (): Document => {
  const document = fromMarkdown("## Scope\n\nBody words.");
  document.package.numbering = {
    abstractNums: [
      {
        abstractNumId: 5,
        multiLevelType: "multilevel",
        levels: [
          {
            ilvl: 0,
            start: 1,
            numFmt: "decimal",
            lvlText: "%1.",
            suffix: "space",
            pPr: { indentLeft: 0, indentFirstLine: 0 },
          },
        ],
      },
    ],
    nums: [{ numId: 5, abstractNumId: 5 }],
  };
  const heading = document.package.styles?.styles.find(
    ({ styleId }) => styleId === "Heading2",
  );
  if (heading === undefined) throw new Error("Heading2 is missing");
  heading.pPr = {
    ...heading.pPr,
    numPr: paragraphNumberingFromSlots({ numId: 5, ilvl: 0 }),
  };
  return document;
};

const open = async (bytes: ArrayBuffer) =>
  FolioDocxReviewer.fromBuffer(bytes, {
    author: "Editor",
    session: "canonical",
  });

const numbering = (document: Document, text: string) => {
  const paragraph = document.package.document.content.find(
    (block): block is Paragraph =>
      block.type === "paragraph" &&
      block.content.some(
        (node) =>
          node.type === "run" &&
          node.content.some(
            (leaf) => leaf.type === "text" && leaf.text.includes(text),
          ),
      ),
  );
  if (paragraph === undefined) throw new Error(`no paragraph holds ${text}`);
  return {
    styleId: paragraph.formatting?.styleId,
    fromStyle: paragraph.formatting?.numPrFromStyle,
    listed: paragraph.listRendering !== undefined,
  };
};

const UNNUMBERED = { styleId: undefined, fromStyle: undefined, listed: false };

test.each(["direct", "tracked-changes"] as const)(
  "clearing a numbered style drops its numbering (%s)",
  async (mode) => {
    const reviewer = await open(await createDocx(source()));
    const [heading] = reviewer.snapshot().blocks;
    if (heading === undefined) throw new Error("the fixture has no block");
    expect(numbering(reviewer.toDocument(), "Scope").listed).toBe(true);
    const result = reviewer.applyDocumentOperations(
      {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode,
        operations: [
          {
            id: "clear",
            type: "setBlockParagraphProperties",
            blockId: heading.id,
            properties: { styleId: null },
          },
          {
            id: "insert",
            type: "insertAfterBlock",
            blockId: heading.id,
            text: "Fresh words.",
            styleId: null,
          },
        ],
      },
      { undefinedReferences: "refuse" },
    );
    expect(result.applied.map(({ id }) => id).toSorted()).toEqual([
      "clear",
      "insert",
    ]);
    const live =
      mode === "direct"
        ? reviewer.toDocument()
        : reviewer.toResolvedDocument("accept");
    const reopened = await open(await reviewer.toBuffer());
    const saved =
      mode === "direct"
        ? reopened.toDocument()
        : reopened.toResolvedDocument("accept");
    for (const text of ["Scope", "Fresh words."]) {
      expect(numbering(live, text)).toEqual(UNNUMBERED);
      expect(numbering(saved, text)).toEqual(UNNUMBERED);
    }
  },
);
