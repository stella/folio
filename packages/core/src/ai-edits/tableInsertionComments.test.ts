/**
 * A row or column added to a table a comment runs across lies inside that
 * comment, even when the same batch deletes the paragraph that held the
 * comment's end. Operations run from the end of the document backwards, so
 * directly the deletion ran first and moved the end into the table's last
 * cell: the new row after it, and the new column's last cell, fell outside.
 * Tracked, the deleted paragraph keeps the end until accepted, and the new
 * cells stayed inside.
 */

import { describe, expect, test } from "bun:test";

import { paragraph, run, table } from "../docx/server/build";
import { createDocx } from "../docx/rezip";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import type { Document } from "../types/document";
import { FolioDocxReviewer } from "./headless";

const MODES = ["direct", "tracked-changes"] as const;

const buildDocument = (): Document => ({
  package: {
    document: {
      comments: [
        {
          id: 7,
          author: "Earlier Reviewer",
          content: [paragraph("Check the whole schedule.")],
        },
      ],
      content: [
        paragraph([{ type: "commentRangeStart", id: 7 }, run("The schedule below is binding.")]),
        table({ header: ["Item", "Price"], rows: [["Gadget", "20"]] }),
        paragraph([
          run("Prices exclude taxes."),
          { type: "commentRangeEnd", id: 7 },
          { type: "commentReference", id: 7 },
        ]),
        paragraph("Signed by both parties."),
      ],
    },
  },
});

type Insert =
  | { type: "insertTableRow"; cellTexts: string[] }
  | { type: "insertTableColumn"; cellTexts: string[] };

const anchoredAfterAccept = async (
  mode: (typeof MODES)[number],
  insert: Insert,
): Promise<string | undefined> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(await createDocx(buildDocument()), {
    author: "Editor",
  });
  const blocks = reviewer.getContent();
  const idOf = (text: string) => {
    const id = blocks.find((block) => block.text === text)?.id;
    if (id === undefined) throw new Error(`fixture must expose "${text}"`);
    return id;
  };
  const result = reviewer.applyDocumentOperations({
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    mode,
    operations: [
      { id: "delete", type: "deleteBlock", blockId: idOf("Prices exclude taxes.") },
      { id: "insert", ...insert, blockId: idOf("20"), position: "after" },
    ],
  });
  expect(result.skipped).toEqual([]);
  if (mode !== "direct") reviewer.acceptAll();
  const reopened = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
  return reopened.getComments()[0]?.anchoredText;
};

describe("a table insertion inside a comment whose end paragraph the batch deletes", () => {
  test.each(MODES)("a new last row stays inside the comment, %s", async (mode) => {
    expect(
      await anchoredAfterAccept(mode, { type: "insertTableRow", cellTexts: ["Widget", "10"] }),
    ).toBe("The schedule below is binding.ItemPriceGadget20Widget10");
  });

  test.each(MODES)("a new last column stays inside the comment, %s", async (mode) => {
    expect(
      await anchoredAfterAccept(mode, { type: "insertTableColumn", cellTexts: ["Tax", "5"] }),
    ).toBe("The schedule below is binding.ItemPriceTaxGadget205");
  });
});
