import { expect, test } from "bun:test";

import { paragraphsDocx, textRun } from "../__tests__/operationBatchDocuments";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { FolioDocxReviewer } from "./headless";

const input = "Lead.";
const source = () => paragraphsDocx([[textRun(input)]]);
const blockId = "10000000";

const edit = (text: string) =>
  ({
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    atomic: true,
    mode: "tracked-changes",
    operations: [
      {
        id: "draft",
        type: "insertAfterBlock",
        blockId,
        text,
        textFormat: "markdown",
      },
    ],
  }) as const;

for (const session of ["canonical"] as const) {
  test(`${session} reviews Markdown headings and inline formatting after a save`, async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await source(), {
      session,
    });
    const result = reviewer.applyDocumentOperations(
      edit("## Terms\n\n**Bold**, *italic*, ~~strike~~ and `code`."),
      { undefinedReferences: "refuse" },
    );
    expect(result.skipped).toEqual([]);
    expect(result.applied.map(({ id }) => id)).toEqual(["draft"]);
    const bytes = await reviewer.toBuffer();

    const accepted = await FolioDocxReviewer.fromBuffer(bytes, { session });
    expect(accepted.acceptAll().skipped).toEqual([]);
    expect(
      accepted.getContent().map(({ text, kind }) => ({ text, kind })),
    ).toEqual([
      { text: input, kind: "paragraph" },
      { text: "Terms", kind: "heading" },
      { text: "Bold, italic, strike and code.", kind: "paragraph" },
    ]);
    const body = accepted.toDocument().package.document.content.at(-1);
    if (body?.type !== "paragraph") throw new Error("Expected a paragraph");
    const formatting = body.content.flatMap((item) =>
      item.type === "run" ? [item.formatting] : [],
    );
    expect(formatting.some((value) => value?.bold === true)).toBe(true);
    expect(formatting.some((value) => value?.italic === true)).toBe(true);
    expect(formatting.some((value) => value?.strike === true)).toBe(true);
    expect(
      formatting.some((value) => value?.fontFamily?.ascii === "Courier New"),
    ).toBe(true);

    const rejected = await FolioDocxReviewer.fromBuffer(bytes, { session });
    expect(rejected.rejectAll().skipped).toEqual([]);
    expect(rejected.getContent().map(({ text }) => text)).toEqual([input]);
  });

  test(`${session} accepts a flat Markdown list as one list instance`, async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await source(), {
      session,
    });
    const result = reviewer.applyDocumentOperations(
      edit("1. First\n2. Second"),
      { undefinedReferences: "refuse" },
    );
    expect(result.skipped).toEqual([]);
    const reopened = await FolioDocxReviewer.fromBuffer(
      await reviewer.toBuffer(),
      { session },
    );
    expect(reopened.acceptAll().skipped).toEqual([]);
    const list = reopened
      .getContent()
      .filter(({ kind }) => kind === "listItem");
    expect(list.map(({ text }) => text)).toEqual(["First", "Second"]);
    expect(list.at(0)?.statedNumbering).toEqual(list.at(1)?.statedNumbering);
  });

  test(`${session} refuses unsupported Markdown without inserting earlier paragraphs`, async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await source(), {
      session,
    });
    const result = reviewer.applyDocumentOperations(
      edit("Supported paragraph.\n\n![Image](https://example.com/image.png)"),
      { undefinedReferences: "refuse" },
    );
    expect(result.applied).toEqual([]);
    expect(result.skipped).toEqual([
      {
        id: "draft",
        reason: "unsupportedBlock",
        message: expect.stringMatching(/image/u),
      },
    ]);
    expect(reviewer.getContent().map(({ text }) => text)).toEqual([input]);
  });
}

const tableEdit = (text: string) =>
  ({
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    atomic: true,
    mode: "tracked-changes",
    operations: [
      {
        id: "draft-table",
        type: "insertTable",
        blockId,
        text,
        textFormat: "markdown",
      },
    ],
  }) as const;

const table = "| Term | Value |\n| --- | --- |\n| Goods | Paid |";

for (const session of ["canonical"] as const) {
  test(`${session} accepts and rejects a tracked Markdown table after reopening`, async () => {
    const document = await paragraphsDocx([
      [textRun(input)],
      [textRun("End.")],
    ]);
    const reviewer = await FolioDocxReviewer.fromBuffer(document, { session });
    const result = reviewer.applyDocumentOperations(tableEdit(table), {
      undefinedReferences: "refuse",
    });
    expect(result.skipped).toEqual([]);
    expect(result.receipts.at(0)).toMatchObject({
      affected: [{ type: "insertion", content: "table" }],
    });
    const bytes = await reviewer.toBuffer();
    const accepted = await FolioDocxReviewer.fromBuffer(bytes, { session });
    expect(accepted.acceptAll().skipped).toEqual([]);
    expect(
      accepted.toDocument().package.document.content.map(({ type }) => type),
    ).toEqual(["paragraph", "table", "paragraph"]);
    expect(accepted.getContent().map(({ text }) => text)).toEqual([
      input,
      "Term",
      "Value",
      "Goods",
      "Paid",
      "End.",
    ]);
    const rejected = await FolioDocxReviewer.fromBuffer(bytes, { session });
    expect(rejected.rejectAll().skipped).toEqual([]);
    expect(rejected.getContent().map(({ text }) => text)).toEqual([
      input,
      "End.",
    ]);
  });
}
