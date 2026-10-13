import { expect, test } from "bun:test";
import { paragraph } from "../docx/server/build";
import { createDocx } from "../docx/rezip";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";

test("paragraph merges use the boundary after batch annotations", async () => {
  for (const mode of ["direct", "tracked-changes", "suggested"] as const) {
    for (const quote of ["First", "part", "end"] as const) {
      const source = createEmptyDocument();
      source.package.document.content = [
        paragraph("First part end"),
        paragraph("Second"),
        paragraph("Third"),
      ];
      const reviewer = await FolioDocxReviewer.fromBuffer(await createDocx(source));
      const first = reviewer.getContent().find(({ text }) => text === "First part end");
      if (!first) throw new Error("Expected first paragraph");
      const result = reviewer.applyDocumentOperations(
        {
          version: 1,
          mode,
          operations: [
            {
              id: "annotate",
              type: "commentOnBlock",
              blockId: first.id,
              quote,
              comment: { text: "Comment" },
            },
            { id: "merge", type: "mergeBlockWithNext", blockId: first.id, separator: " " },
          ],
        },
        { undefinedReferences: "refuse" },
      );
      expect(result.skipped).toEqual([]);
      expect(result.applied.map(({ id }) => id).toSorted()).toEqual(["annotate", "merge"]);
      reviewer.acceptAll();
      const reopened = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
      expect(reopened.getContent().map(({ text }) => text)).toEqual([
        "First part end Second",
        "Third",
      ]);
      expect(reopened.getComments().map(({ anchoredText }) => anchoredText)).toEqual([quote]);
    }
  }
});
