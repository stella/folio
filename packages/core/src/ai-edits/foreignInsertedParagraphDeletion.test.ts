/**
 * Deleting a paragraph another author inserted records the deletion on its
 * mark beside their insertion, on both sessions: the saved mark holds
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
  session: "legacy" | "canonical",
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

for (const session of ["legacy", "canonical"] as const) {
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
