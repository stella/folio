/**
 * A paragraph whose words and mark are deleted, followed by two inserted body
 * paragraphs and then a heading numbered through its style: accepting every
 * change keeps the inserted paragraphs body text. This is the written redline
 * of the consumer compare round trip on the style-numbered fixture.
 */

import { describe, expect, test } from "bun:test";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { createDocx } from "../docx/rezip";
import type { Paragraph } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";

const INFO = { author: "Comparison", date: "2026-11-15T00:00:00Z" } as const;

const text = (words: string) => ({
  type: "run" as const,
  content: [{ type: "text" as const, text: words }],
});

const kept = (paraId: string, words: string, styleId?: string): Paragraph => ({
  type: "paragraph",
  paraId,
  ...(styleId === undefined ? {} : { formatting: { styleId } }),
  content: [text(words)],
});

const BODY: Paragraph[] = [
  kept("10000001", "Agreement", "Heading2"),
  kept("10000002", "The Supplier delivers the goods."),
  {
    type: "paragraph",
    paraId: "10000003",
    pPrMark: { kind: "del", info: { id: 1, ...INFO } },
    content: [
      {
        type: "deletion",
        info: { id: 2, ...INFO },
        content: [text("The Buyer pays on delivery.")],
      },
    ],
  },
  {
    type: "paragraph",
    paraId: "10000004",
    pPrMark: { kind: "ins", info: { id: 3, ...INFO } },
    content: [
      {
        type: "insertion",
        info: { id: 4, ...INFO },
        content: [text("Buyer buyer warranty payment period.")],
      },
    ],
  },
  {
    type: "paragraph",
    paraId: "10000005",
    pPrMark: { kind: "ins", info: { id: 5, ...INFO } },
    content: [
      {
        type: "insertion",
        info: { id: 6, ...INFO },
        content: [text("Notice warranty invoice notice buyer.")],
      },
    ],
  },
  kept("10000006", "Payment", "Heading2"),
  kept("10000007", "Closing paragraph."),
];

const redline = (): Promise<ArrayBuffer> => {
  const document = createEmptyDocument();
  document.package.styles = {
    styles: [
      { type: "paragraph", styleId: "Normal", name: "Normal", default: true },
      {
        type: "paragraph",
        styleId: "Heading2",
        name: "heading 2",
        basedOn: "Normal",
        pPr: {
          outlineLevel: 1,
          numPr: { kind: "reference", numId: 5, ilvl: 0 },
        },
      },
    ],
  };
  document.package.numbering = {
    abstractNums: [
      {
        abstractNumId: 5,
        multiLevelType: "multilevel",
        levels: [
          { ilvl: 0, start: 1, numFmt: "decimal", lvlText: "%1." },
          { ilvl: 1, start: 1, numFmt: "decimal", lvlText: "%1.%2." },
        ],
      },
    ],
    nums: [{ numId: 5, abstractNumId: 5 }],
  };
  document.package.document.content = BODY;
  return createDocx(document);
};

const show = (reviewer: FolioDocxReviewer) =>
  reviewer.getContent().map((block) => JSON.stringify(block).slice(0, 400));

describe("L3E scratch", () => {
  for (const session of ["legacy", "canonical"] as const)
    for (const mode of ["accept", "reject"] as const)
      test(`${session} ${mode}`, async () => {
        const reviewer = await FolioDocxReviewer.fromBuffer(
          (await redline()).slice(0),
          { session },
        );
        console.log("L3E", session, mode, "before", show(reviewer));
        if (mode === "accept") reviewer.acceptAll();
        else reviewer.rejectAll();
        console.log("L3E", session, mode, "resolved", show(reviewer));
        const reopened = await FolioDocxReviewer.fromBuffer(
          await reviewer.toBuffer(),
          { session },
        );
        console.log("L3E", session, mode, "reopened", show(reopened));
      });
});

describe("inserted body paragraphs before a style-numbered heading", () => {
  test("stay body text once every change is accepted", async () => {
    const buffer = await redline();
    for (const session of ["legacy", "canonical"] as const) {
      const reviewer = await FolioDocxReviewer.fromBuffer(buffer.slice(0), {
        session,
      });
      reviewer.acceptAll();
      const reopened = await FolioDocxReviewer.fromBuffer(
        await reviewer.toBuffer(),
        { session },
      );
      expect({
        session,
        blocks: reopened
          .getContent()
          .map((block) =>
            "kind" in block && "text" in block
              ? `${String(block.kind)}: ${String(block.text)}`
              : "other",
          ),
      }).toEqual({
        session,
        blocks: [
          "heading: Agreement",
          "paragraph: The Supplier delivers the goods.",
          "paragraph: Buyer buyer warranty payment period.",
          "paragraph: Notice warranty invoice notice buyer.",
          "heading: Payment",
          "paragraph: Closing paragraph.",
        ],
      });
    }
  });
});
