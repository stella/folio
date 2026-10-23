import { describe, expect, test } from "bun:test";

import { fromMarkdown } from ".";
import { FolioDocxReviewer } from "../ai-edits/headless";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { createDocx } from "../docx/rezip";
import { docxToMarkdown } from "../docx/server/docxToMarkdown";

describe("clean Markdown of a pending numbering removal", () => {
  test("reads the paragraph as accepting leaves it, as content readers do", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(
      await createDocx(
        fromMarkdown(
          ["1. Deposit on signature", "2. Balance on delivery"].join("\n"),
        ),
      ),
      { author: "Reviewer" },
    );
    const balance = reviewer
      .getContent()
      .find(({ text }) => text === "Balance on delivery");
    const result = reviewer.applyDocumentOperations(
      {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "tracked-changes",
        operations: [
          {
            id: "unnumber",
            type: "setBlockParagraphProperties",
            blockId: balance?.id ?? "",
            properties: { numbering: { kind: "inherit" } },
          },
        ],
      },
      { undefinedReferences: "refuse" },
    );
    expect(result.skipped).toEqual([]);
    const saved = await reviewer.toBuffer();

    const accepted = await FolioDocxReviewer.fromBuffer(saved, {
      author: "Reviewer",
    });
    accepted.acceptAll();
    expect(
      accepted.getContent().find(({ text }) => text === "Balance on delivery"),
    ).toMatchObject({ kind: "paragraph" });

    const clean = await docxToMarkdown(saved, {
      annotations: "strip",
      trackedChanges: "clean",
      comments: "strip",
    });
    expect(clean).toContain("1. Deposit on signature");
    expect(clean).not.toContain("2. Balance on delivery");
    expect(clean).toContain("Balance on delivery");
  });
});
