import { expect, test } from "bun:test";
import { createDocx } from "../docx/rezip";
import { parseDocx } from "../docx/parser";
import { paragraph } from "../docx/server/build";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";

test("direct merges retain the ending paragraph mark revision", async () => {
  for (const kind of ["ins", "del", "moveFrom", "moveTo"] as const) {
    const document = createEmptyDocument();
    const ending = paragraph("Second");
    ending.pPrMark = { kind, info: { id: 7, author: "Reviewer", date: "2026-01-01T00:00:00Z" } };
    document.package.document.content = [paragraph("First"), ending, paragraph("Third")];
    const reviewer = await FolioDocxReviewer.fromBuffer(await createDocx(document));
    const first = reviewer.getContent().find(({ text }) => text === "First");
    if (!first) throw new Error("Expected first paragraph");
    const result = reviewer.applyDocumentOperations(
      {
        version: 1,
        mode: "direct",
        operations: [
          { id: "merge", type: "mergeBlockWithNext", blockId: first.id, separator: " " },
        ],
      },
      { undefinedReferences: "refuse" },
    );
    expect(result.skipped).toEqual([]);
    const saved = await parseDocx(await reviewer.toBuffer());
    const merged = saved.package.document.content.at(0);
    expect(merged?.type).toBe("paragraph");
    if (merged?.type !== "paragraph") throw new Error("Expected merged paragraph");
    expect(merged.pPrMark).toEqual(ending.pPrMark);
  }
});
