/**
 * A body paragraph rewritten into two body paragraphs with other wording,
 * the second keeping its id, right before a heading that states its outline
 * level itself. The comparison aligns them by content, so it writes the old
 * paragraph deleted and the two new ones inserted before the heading: the
 * inserted paragraphs are body text as written, and accepting the redline
 * gives the revised version.
 */

import { describe, expect, test } from "bun:test";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { createDocx } from "../docx/rezip";
import type { Paragraph } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { compareDocx, parseComparison, planComparison } from "./compare";

const OPTIONS = {
  author: "compare",
  timestamp: "2026-11-15T00:00:00.000Z",
} as const;

const body = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  textId: paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const heading = (paraId: string, text: string): Paragraph => ({
  ...body(paraId, text),
  formatting: {
    styleId: "Heading2",
    outlineLevel: { kind: "heading", level: 1 },
  },
});

const packageOf = (content: Paragraph[]): Promise<ArrayBuffer> => {
  const document = createEmptyDocument();
  document.package.styles = {
    styles: [
      { type: "paragraph", styleId: "Normal", name: "Normal", default: true },
      {
        type: "paragraph",
        styleId: "Heading2",
        name: "heading 2",
        basedOn: "Normal",
      },
    ],
  };
  document.package.document.content = content;
  return createDocx(document);
};

const BEFORE = [
  body("10000001", "The Supplier delivers the goods."),
  body("10000002", "The Buyer pays on delivery."),
  heading("10000003", "Payment"),
  body("10000004", "Closing paragraph."),
];

const AFTER = [
  body("10000001", "The Supplier delivers the goods."),
  body("10000005", "Buyer buyer warranty payment period."),
  body("10000002", "Notice warranty invoice notice buyer."),
  heading("10000003", "Payment"),
  body("10000004", "Closing paragraph."),
];

/** Each body paragraph: its id, properties, tracked mark and property changes, and its content. */
const layoutOf = (reviewer: FolioDocxReviewer) =>
  reviewer.toDocument().package.document.content.flatMap((block) =>
    block.type === "paragraph"
      ? [
          {
            paraId: block.paraId,
            formatting: block.formatting,
            pPrMark: block.pPrMark,
            propertyChanges: block.propertyChanges,
            content: JSON.stringify(block.content),
          },
        ]
      : [],
  );

describe("a paragraph rewritten into two before a heading stating its level", () => {
  test("is written as body paragraphs and accepts to the revised version", async () => {
    const before = await packageOf(BEFORE);
    const after = await packageOf(AFTER);
    const parsed = await parseComparison(before, after, OPTIONS);
    if (parsed.isErr()) throw parsed.error;
    const planned = planComparison(parsed.value);
    if (planned.isErr()) throw planned.error;
    const operations = planned.value.flatMap(({ plan }) => plan.operations);

    const compared = await compareDocx(before, after, OPTIONS);
    if (compared.isErr()) throw compared.error;
    const written = await FolioDocxReviewer.fromBuffer(compared.value.buffer);
    const layout = layoutOf(written);
    // No paragraph holding only new words states an outline level or a
    // style before any resolution: the revised paragraphs state neither.
    expect({
      operations,
      layout,
      inserted: layout.filter(
        ({ pPrMark, formatting }) =>
          pPrMark?.kind === "ins" &&
          (formatting?.outlineLevel !== undefined ||
            formatting?.styleId !== undefined),
      ),
    }).toMatchObject({ inserted: [] });

    written.acceptAll();
    expect({
      operations,
      layout,
      accepted: written
        .getContent()
        .map((block) =>
          "kind" in block && "text" in block
            ? `${String(block.kind)}: ${String(block.text)}`
            : "other",
        ),
    }).toMatchObject({
      accepted: [
        "paragraph: The Supplier delivers the goods.",
        "paragraph: Buyer buyer warranty payment period.",
        "paragraph: Notice warranty invoice notice buyer.",
        "heading: Payment",
        "paragraph: Closing paragraph.",
      ],
    });
  });
});
