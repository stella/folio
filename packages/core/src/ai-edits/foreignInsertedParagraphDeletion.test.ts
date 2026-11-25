/**
 * Deleting a paragraph another author inserted records the deletion on its
 * mark beside their insertion: the saved mark holds
 * `w:ins` and `w:del`, accepting or rejecting everything removes the
 * paragraph, and each record resolves on its own.
 */
import { expect, test } from "bun:test";

import { textRun } from "../__tests__/operationBatchDocuments";
import { createDocx } from "../docx/rezip";
import { unzipDocx } from "../docx/unzip";
import type { Document } from "../types/document";
import { FolioDocxReviewer } from "./headless";

const table = {
  type: "table" as const,
  rows: [
    {
      type: "tableRow" as const,
      cells: [
        {
          type: "tableCell" as const,
          content: [
            {
              type: "paragraph" as const,
              paraId: "20000000",
              content: [textRun("Cell")],
            },
          ],
        },
      ],
    },
  ],
};

const texts = (document: Document): string[] =>
  document.package.document.content.map((block) =>
    block.type === "paragraph"
      ? block.content
          .flatMap((item) => (item.type === "run" ? item.content : []))
          .map((item) => (item.type === "text" ? item.text : ""))
          .join("")
      : block.type,
  );

const deleted = async (
  session: "canonical",
  beforeTable: boolean,
): Promise<ArrayBuffer> => {
  const docx = await createDocx({
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "10000000",
            content: [textRun("First")],
          },
          {
            type: "paragraph",
            paraId: "10000001",
            content: [
              {
                type: "insertion",
                info: { id: 1, author: "First Reviewer" },
                content: [textRun("Added")],
              },
            ],
            pPrMark: {
              kind: "ins",
              info: { id: 2, author: "First Reviewer" },
            },
          },
          ...(beforeTable ? [table] : []),
          {
            type: "paragraph",
            paraId: "10000002",
            content: [textRun("Final")],
          },
        ],
      },
    },
  });
  const reviewer = await FolioDocxReviewer.fromBuffer(docx, {
    author: "Second Reviewer",
    session,
  });
  const result = reviewer.applyDocumentOperations(
    {
      version: 1,
      mode: "tracked-changes",
      operations: [{ id: "delete", type: "deleteBlock", blockId: "10000001" }],
    },
    { undefinedReferences: "refuse" },
  );
  expect(result.skipped).toEqual([]);
  return await reviewer.toBuffer();
};

for (const session of ["canonical"] as const) {
  for (const beforeTable of [false, true]) {
    const where = beforeTable ? "before a table" : "between paragraphs";
    const kept = beforeTable ? ["First", "table", "Final"] : ["First", "Final"];
    const emptied = beforeTable
      ? ["First", "", "table", "Final"]
      : ["First", "", "Final"];

    test(`${session}: deleting another author's inserted paragraph ${where} records both records on its mark`, async () => {
      const saved = await deleted(session, beforeTable);
      const xml = (await unzipDocx(saved)).documentXml ?? "";
      expect(xml).toMatch(
        /<w:pPr><w:rPr><w:ins w:id="2" w:author="First Reviewer"\/><w:del w:id="\d+" w:author="Second Reviewer"[^>]*\/><\/w:rPr><\/w:pPr>/u,
      );
      const reopened = await FolioDocxReviewer.fromBuffer(saved, { session });
      const marks = reopened
        .getChanges()
        .filter(({ type }) => type.startsWith("paragraphMark"))
        .map(({ type, author }) => ({ type, author }));
      expect(marks).toEqual([
        { type: "paragraphMarkInserted", author: "First Reviewer" },
        { type: "paragraphMarkDeleted", author: "Second Reviewer" },
      ]);
      expect(texts(reopened.toResolvedDocument("accept"))).toEqual(kept);
      expect(texts(reopened.toResolvedDocument("reject"))).toEqual(kept);
    });

    test(`${session}: each record of the mark ${where} resolves on its own`, async () => {
      const saved = await deleted(session, beforeTable);
      const markIds = async () => {
        const reviewer = await FolioDocxReviewer.fromBuffer(saved, {
          session,
        });
        const changes = reviewer.getChanges();
        const id = (type: string) =>
          changes.find((change) => change.type === type)?.id ?? -1;
        return {
          insertion: id("paragraphMarkInserted"),
          deletion: id("paragraphMarkDeleted"),
        };
      };
      const { insertion, deletion } = await markIds();
      const cases = [
        {
          id: insertion,
          accept: true,
          after: emptied,
          left: "paragraphMarkDeleted",
        },
        { id: deletion, accept: true, after: kept, left: undefined },
        { id: insertion, accept: false, after: kept, left: undefined },
        {
          id: deletion,
          accept: false,
          after: emptied,
          left: "paragraphMarkInserted",
        },
      ];
      for (const { id, accept, after, left } of cases) {
        const reviewer = await FolioDocxReviewer.fromBuffer(saved, {
          session,
        });
        expect(
          accept ? reviewer.acceptChange(id) : reviewer.rejectChange(id),
        ).toBe(true);
        expect(texts(reviewer.toDocument())).toEqual(after);
        expect(
          reviewer
            .getChanges()
            .filter(({ type }) => type.startsWith("paragraphMark"))
            .map(({ type }) => type),
        ).toEqual(left === undefined ? [] : [left]);
      }
    });
  }
}

for (const session of ["canonical"] as const) {
  test(`${session}: deleting the final paragraph after another author's added break records its deletion beside the insertion`, async () => {
    const docx = await createDocx({
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              paraId: "10000000",
              content: [textRun("First")],
            },
            {
              type: "paragraph",
              paraId: "10000001",
              content: [textRun("Split head")],
              pPrMark: {
                kind: "ins",
                info: { id: 2, author: "First Reviewer" },
              },
            },
            {
              type: "paragraph",
              paraId: "10000002",
              content: [textRun("tail")],
            },
          ],
        },
      },
    });
    const reviewer = await FolioDocxReviewer.fromBuffer(docx, {
      author: "Second Reviewer",
      session,
    });
    const result = reviewer.applyDocumentOperations(
      {
        version: 1,
        mode: "tracked-changes",
        operations: [
          { id: "delete", type: "deleteBlock", blockId: "10000002" },
        ],
      },
      { undefinedReferences: "refuse" },
    );
    expect(result.skipped).toEqual([]);
    const saved = await reviewer.toBuffer();
    expect((await unzipDocx(saved)).documentXml ?? "").toMatch(
      /<w:rPr><w:ins w:id="2" w:author="First Reviewer"\/><w:del w:id="\d+" w:author="Second Reviewer"[^>]*\/><\/w:rPr>/u,
    );
    const reopened = await FolioDocxReviewer.fromBuffer(saved, { session });
    expect(texts(reopened.toResolvedDocument("accept"))).toEqual([
      "First",
      "Split head",
    ]);
    expect(texts(reopened.toResolvedDocument("reject"))).toEqual([
      "First",
      "Split headtail",
    ]);
  });
}

const takenMark = {
  kind: "ins" as const,
  info: { id: 2, author: "First Reviewer" },
  deletion: { info: { id: 3, author: "Second Reviewer" } },
};

for (const session of ["canonical"] as const) {
  for (const target of ["the paragraph", "the final paragraph"] as const) {
    test(`${session}: deleting ${target} again keeps both records of a mark another deletion already took`, async () => {
      const docx = await createDocx({
        package: {
          document: {
            content: [
              {
                type: "paragraph",
                paraId: "10000000",
                content: [textRun("First")],
              },
              {
                type: "paragraph",
                paraId: "10000001",
                content: [textRun("Added")],
                pPrMark: takenMark,
              },
              {
                type: "paragraph",
                paraId: "10000002",
                content: [textRun("tail")],
              },
            ],
          },
        },
      });
      const reviewer = await FolioDocxReviewer.fromBuffer(docx, {
        author: "Third Reviewer",
        session,
      });
      reviewer.applyDocumentOperations(
        {
          version: 1,
          mode: "tracked-changes",
          operations: [
            {
              id: "delete",
              type: "deleteBlock",
              blockId: target === "the paragraph" ? "10000001" : "10000002",
            },
          ],
        },
        { undefinedReferences: "refuse" },
      );
      const xml =
        (await unzipDocx(await reviewer.toBuffer())).documentXml ?? "";
      expect(xml).toMatch(
        /<w:rPr><w:ins w:id="2" w:author="First Reviewer"\/><w:del w:id="3" w:author="Second Reviewer"\/><\/w:rPr>/u,
      );
    });
  }
}
