/**
 * `formattingScope` on multiline insertions (issue #845). One `insertAfterBlock`
 * whose `text` splits into several paragraphs formats only the first by default,
 * which is right for a heading followed by its body but wrong for several list
 * items. `"allParagraphs"` numbers every paragraph, and the membership must
 * survive a DOCX save and reopen, not just the in-memory snapshot.
 */

import { describe, expect, test } from "bun:test";

import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import type { FolioDocumentOperation } from "../document-operations";
import { FolioDocxReviewer } from "./headless";
import type { FolioAIBlock, FolioAIInsertFormattingScope } from "./types";

type InsertExtras = Partial<
  Pick<
    Extract<FolioDocumentOperation, { type: "insertAfterBlock" }>,
    "formattingScope" | "listLevel" | "numbering" | "styleId"
  >
>;

const numberedListDocx = async (): Promise<{ docx: Uint8Array; numId: number }> => {
  const model = fromMarkdown("1. Alpha\n2. Beta\n\nTail.");
  const beta = model.package.document.content.at(1);
  const numId = beta?.type === "paragraph" ? beta.formatting?.numPr?.numId : undefined;
  if (numId === undefined) {
    throw new Error("fixture must number its second paragraph");
  }
  const { docx } = await ensureParaIds(await createDocx(model));
  return { docx, numId };
};

type InsertedOutline = { text: string; kind: string; marker: string | null }[];

type InsertAfterOptions = {
  anchorText: string;
  text: string;
  extras: InsertExtras;
};

/** Apply one tracked insertion, save, reopen, and return the reopened blocks. */
const insertAfter = async ({
  anchorText,
  text,
  extras,
}: InsertAfterOptions): Promise<FolioAIBlock[]> => {
  const { docx } = await numberedListDocx();
  const reviewer = await FolioDocxReviewer.fromBuffer(docx, { author: "Agent" });
  const anchor = reviewer.getContent().find((block) => block.text === anchorText);
  if (!anchor) {
    throw new Error(`fixture must expose the ${anchorText} anchor`);
  }
  const result = reviewer.applyDocumentOperations({
    version: 1,
    mode: "tracked-changes",
    operations: [{ id: "insert", type: "insertAfterBlock", blockId: anchor.id, text, ...extras }],
  });
  expect(result.status).toBe("committed");
  expect(result.issues).toEqual([]);

  const reopened = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
  return reopened.getContent();
};

const outlineOf = (blocks: FolioAIBlock[]): InsertedOutline =>
  blocks.map((block) => ({
    text: block.text,
    kind: block.kind,
    marker: block.displayLabel ?? null,
  }));

const insertAfterBeta = async (text: string, extras: InsertExtras): Promise<InsertedOutline> =>
  outlineOf(await insertAfter({ anchorText: "Beta", text, extras }));

const THREE_ITEMS = "Gamma\nDelta\nEpsilon";

const listItem = (text: string, marker: string) => ({ text, kind: "listItem", marker });
const bodyParagraph = (text: string) => ({ text, kind: "paragraph", marker: null });

const LIST_HEAD = [listItem("Alpha", "1."), listItem("Beta", "2.")];
const TAIL = [bodyParagraph("Tail.")];

describe("insertAfterBlock formattingScope", () => {
  const allParagraphVariants: { name: string; extras: (numId: number) => InsertExtras }[] = [
    { name: "inherited numbering", extras: () => ({}) },
    { name: "explicit listLevel", extras: () => ({ listLevel: 0 }) },
    { name: "explicit numbering", extras: (numId) => ({ numbering: { numId, level: 0 } }) },
  ];

  for (const { name, extras } of allParagraphVariants) {
    test(`"allParagraphs" numbers every split paragraph after save and reopen: ${name}`, async () => {
      const { numId } = await numberedListDocx();
      const outline = await insertAfterBeta(THREE_ITEMS, {
        formattingScope: "allParagraphs",
        ...extras(numId),
      });
      expect(outline).toEqual([
        ...LIST_HEAD,
        listItem("Gamma", "3."),
        listItem("Delta", "4."),
        listItem("Epsilon", "5."),
        ...TAIL,
      ]);
    });
  }

  const firstParagraphScopes: (FolioAIInsertFormattingScope | undefined)[] = [
    undefined,
    "firstParagraph",
  ];

  for (const formattingScope of firstParagraphScopes) {
    test(`${formattingScope ?? "the default"} scope numbers only the first split paragraph`, async () => {
      const { numId } = await numberedListDocx();
      const outline = await insertAfterBeta(THREE_ITEMS, {
        ...(formattingScope !== undefined && { formattingScope }),
        numbering: { numId, level: 0 },
      });
      expect(outline).toEqual([
        ...LIST_HEAD,
        listItem("Gamma", "3."),
        bodyParagraph("Delta"),
        bodyParagraph("Epsilon"),
        ...TAIL,
      ]);
    });
  }

  test("a single paragraph continues the list in either scope", async () => {
    for (const formattingScope of ["firstParagraph", "allParagraphs"] as const) {
      const outline = await insertAfterBeta("Gamma", { formattingScope });
      expect(outline).toEqual([...LIST_HEAD, listItem("Gamma", "3."), ...TAIL]);
    }
  });

  test("the default scope keeps a heading followed by its body", async () => {
    const blocks = await insertAfter({
      anchorText: "Tail.",
      text: "Heading\nBody.",
      extras: { styleId: "Heading1" },
    });
    expect(blocks.map((block) => [block.text, block.styleId ?? null])).toEqual([
      ["Alpha", null],
      ["Beta", null],
      ["Tail.", null],
      ["Heading", "Heading1"],
      ["Body.", null],
    ]);
    expect(outlineOf(blocks).slice(3)).toEqual([
      { text: "Heading", kind: "heading", marker: "Heading1" },
      bodyParagraph("Body."),
    ]);
  });
});
