/**
 * Two headings rewritten into body paragraphs with other wording keep their
 * paragraph ids. The comparison aligns them by content, so it writes the
 * headings deleted and the body paragraphs inserted: the written redline
 * shows the inserted paragraphs as body text, and accepting it gives the
 * revised version's body paragraphs, rejecting it the original headings.
 */

import { describe, expect, test } from "bun:test";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { createDocx } from "../docx/rezip";
import type { BlockContent, Paragraph } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { compareDocx } from "./compare";

const OPTIONS = {
  author: "compare",
  timestamp: "2026-11-15T00:00:00.000Z",
} as const;

const paragraph = (
  paraId: string,
  text: string,
  styleId?: string,
): Paragraph => ({
  type: "paragraph",
  paraId,
  textId: paraId,
  ...(styleId === undefined ? {} : { formatting: { styleId } }),
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const packageOf = (content: Paragraph[]): Promise<ArrayBuffer> => {
  const document = createEmptyDocument();
  document.package.styles = {
    styles: [
      { type: "paragraph", styleId: "Normal", name: "Normal", default: true },
      {
        type: "paragraph",
        styleId: "Heading1",
        name: "heading 1",
        basedOn: "Normal",
      },
    ],
  };
  document.package.document.content = content;
  return createDocx(document);
};

const BEFORE = [
  paragraph("10000001", "The Supplier delivers the goods."),
  paragraph("10000002", "Warranty", "Heading1"),
  paragraph("10000003", "Notice", "Heading1"),
  paragraph("10000004", "Closing paragraph."),
];

const AFTER = [
  paragraph("10000001", "The Supplier delivers the goods."),
  paragraph("10000002", "Buyer buyer warranty payment period."),
  paragraph("10000003", "Notice warranty invoice notice buyer."),
  paragraph("10000004", "Closing paragraph."),
];

/** Each tracked wrapper's kind and the words it holds, in order. */
const piecesOf = (node: unknown, wrapper: string): [string, string][] => {
  if (typeof node !== "object" || node === null) return [];
  const record = node as Record<string, unknown>;
  const kind =
    record["type"] === "insertion" || record["type"] === "deletion"
      ? String(record["type"])
      : wrapper;
  if (record["type"] === "text" && typeof record["text"] === "string")
    return [[kind, record["text"]]];
  const content = record["content"];
  return Array.isArray(content)
    ? content.flatMap((child) => piecesOf(child, kind))
    : [];
};

/** Each body paragraph: its style, its mark's tracked kind and its words. */
const layoutOf = (blocks: readonly BlockContent[]) =>
  blocks.flatMap((block) =>
    block.type === "paragraph"
      ? [
          {
            styleId: block.formatting?.styleId ?? null,
            mark: block.pPrMark?.kind ?? null,
            propertyChanges: block.propertyChanges?.length ?? 0,
            pieces: piecesOf(block, "kept"),
          },
        ]
      : [],
  );

const wordsAndStyles = (blocks: readonly BlockContent[]) =>
  layoutOf(blocks).map(({ styleId, pieces }) => ({
    styleId,
    text: pieces.map(([, text]) => text).join(""),
  }));

describe("headings rewritten into body paragraphs", () => {
  test("are written as body paragraphs and accept to the revised version", async () => {
    const compared = await compareDocx(
      await packageOf(BEFORE),
      await packageOf(AFTER),
      OPTIONS,
    );
    if (compared.isErr()) throw compared.error;
    const written = await FolioDocxReviewer.fromBuffer(compared.value.buffer);
    const layout = layoutOf(written.toDocument().package.document.content);
    // No paragraph holding only the new words is written with the heading
    // style, before any resolution.
    expect({
      layout,
      inserted: layout.filter(
        ({ styleId, pieces }) =>
          styleId === "Heading1" &&
          pieces.length > 0 &&
          pieces.every(([kind]) => kind === "insertion"),
      ),
    }).toMatchObject({ inserted: [] });

    const accepted = await FolioDocxReviewer.fromBuffer(compared.value.buffer);
    accepted.acceptAll();
    expect({
      layout,
      accepted: wordsAndStyles(accepted.toDocument().package.document.content),
    }).toMatchObject({
      accepted: AFTER.map((block) => ({
        styleId: null,
        text: piecesOf(block, "kept")
          .map(([, text]) => text)
          .join(""),
      })),
    });
  });
});
