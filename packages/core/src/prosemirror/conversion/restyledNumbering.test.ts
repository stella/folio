import { describe, expect, test } from "bun:test";

import { paragraphNumberingReference } from "@stll/docx-core/model";

import { FolioDocxReviewer } from "../../ai-edits/headless";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../../document-operations";
import { ensureParaIds } from "../../docx/ensureParaIds";
import { createDocx } from "../../docx/rezip";
import { fromMarkdown } from "../../markdown";
import type { Paragraph } from "../../types/document";

/** "## Scope" numbered through its `Heading2` style. */
const styleNumberedDocx = async (): Promise<ArrayBuffer> => {
  const document = fromMarkdown(
    ["# Agreement", "## Scope", "Body."].join("\n\n"),
  );
  document.package.numbering = {
    abstractNums: [
      {
        abstractNumId: 5,
        levels: [{ ilvl: 0, numFmt: "decimal", lvlText: "%1." }],
      },
    ],
    nums: [{ numId: 5, abstractNumId: 5 }],
  };
  for (const style of document.package.styles?.styles ?? []) {
    if (style.styleId === "Heading2") {
      style.pPr = {
        ...style.pPr,
        numPr: paragraphNumberingReference({ numId: 5, ilvl: 0 }),
      };
    }
  }
  const bytes = (
    await ensureParaIds(new Uint8Array(await createDocx(document)))
  ).docx;
  return bytes.buffer.slice(
    bytes.byteOffset,
    bytes.byteOffset + bytes.byteLength,
  ) as ArrayBuffer;
};

const scopeOf = (paragraphs: readonly unknown[]): Paragraph | undefined =>
  paragraphs.find(
    (block): block is Paragraph =>
      typeof block === "object" &&
      block !== null &&
      "type" in block &&
      block.type === "paragraph" &&
      JSON.stringify(block).includes('"Scope"'),
  );

describe("a restyled paragraph's inherited numbering", () => {
  for (const mode of ["direct", "tracked-changes"] as const) {
    test(`leaves with the style it came from (${mode})`, async () => {
      const reviewer = await FolioDocxReviewer.fromBuffer(
        await styleNumberedDocx(),
        { author: "Reviewer" },
      );
      const scope = reviewer.getContent().find(({ text }) => text === "Scope");
      expect(scope?.kind).toBe("heading");
      const result = reviewer.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode,
          operations: [
            {
              id: "clear",
              type: "setBlockParagraphProperties",
              blockId: scope?.id ?? "",
              properties: { styleId: null },
            },
          ],
        },
        { undefinedReferences: "refuse" },
      );
      expect(result.skipped).toEqual([]);

      const reopened = await FolioDocxReviewer.fromBuffer(
        await reviewer.toBuffer(),
        { author: "Reviewer" },
      );
      const live =
        mode === "direct"
          ? reviewer.toDocument()
          : reviewer.toResolvedDocument("accept");
      const saved =
        mode === "direct"
          ? reopened.toDocument()
          : reopened.toResolvedDocument("accept");
      for (const document of [live, saved]) {
        const paragraph = scopeOf(document.package.document.content);
        expect(paragraph?.formatting?.numPrFromStyle).toBeUndefined();
        expect(paragraph?.listRendering).toBeUndefined();
      }
    });
  }
});
