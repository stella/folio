/**
 * An operation whose block an earlier operation of the same batch removed is
 * refused with a reason (`overlappingOperation`), not applied to whatever took
 * the block's place and not thrown out of the batch.
 */

import { describe, expect, test } from "bun:test";

import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { paragraph } from "../docx/server/build";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";

const buildReviewer = async (): Promise<FolioDocxReviewer> => {
  const document = createEmptyDocument();
  document.package.document.content = [
    paragraph("First clause."),
    paragraph("Second clause."),
    paragraph("Third clause."),
  ];
  const { docx } = await ensureParaIds(new Uint8Array(await createDocx(document)));
  return FolioDocxReviewer.fromBuffer(
    docx.buffer.slice(docx.byteOffset, docx.byteOffset + docx.byteLength) as ArrayBuffer,
  );
};

describe("a batch that removes a block and then formats it", () => {
  test("refuses the formatting and leaves the next paragraph alone", async () => {
    const reviewer = await buildReviewer();
    const second = reviewer.getContent().find(({ text }) => text === "Second clause.");
    if (!second) {
      throw new Error("the fixture paragraph is missing");
    }
    const result = reviewer.applyDocumentOperations({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "direct",
      operations: [
        { id: "delete", type: "deleteBlock", blockId: second.id },
        {
          id: "center",
          type: "setBlockParagraphProperties",
          blockId: second.id,
          properties: { alignment: "center" },
        },
      ],
    });

    expect(result.applied.map(({ id }) => id)).toEqual(["delete"]);
    expect(result.skipped).toEqual([
      {
        id: "center",
        reason: "overlappingOperation",
        message: 'operation "delete", earlier in this batch, already claims its target.',
      },
    ]);
    expect(reviewer.getContent().map(({ text }) => text)).toEqual([
      "First clause.",
      "Third clause.",
    ]);
    expect(reviewer.getContent().some(({ directAlignment }) => directAlignment !== undefined)).toBe(
      false,
    );
    await reviewer.toBuffer();
  });
});
