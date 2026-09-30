import { expect, test } from "bun:test";
import type { Document } from "@stll/docx-core/model";

import { reviewDifferences } from "./reviewOracle";

const fixture = (): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: "00000001",
          content: [
            {
              type: "insertion",
              info: { id: 41, author: "Reviewer", date: "2026-05-06T07:08:09Z" },
              content: [
                {
                  type: "run",
                  content: [{ type: "text", text: "Tracked" }],
                },
              ],
            },
          ],
        },
      ],
    },
  },
});

test("the oracle detects lost text and revision identities", () => {
  const original = fixture();
  expect(reviewDifferences(original, structuredClone(original))).toEqual({
    messages: [],
    omitted: 0,
  });
  const changed = structuredClone(original);
  const paragraph = changed.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") throw new Error("Expected paragraph");
  const insertion = paragraph.content.at(0);
  if (insertion?.type !== "insertion") throw new Error("Expected insertion");
  insertion.info.id = 42;
  expect(reviewDifferences(original, changed).messages).not.toEqual([]);
  insertion.info.id = 41;
  insertion.content = [];
  expect(reviewDifferences(original, changed).messages).not.toEqual([]);
});

test("the oracle detects a different paragraph survivor", () => {
  const original = fixture();
  const changed = structuredClone(original);
  const paragraph = changed.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") throw new Error("Expected paragraph");
  paragraph.paraId = "00000002";
  expect(reviewDifferences(original, changed).messages).not.toEqual([]);
});

test("empty formatting is neutral while explicit false remains authored", () => {
  const original = fixture();
  original.package.document.content.push({
    type: "table",
    rows: [
      {
        type: "tableRow",
        cells: [
          { type: "tableCell", content: [{ type: "paragraph", paraId: "00000002", content: [] }] },
        ],
      },
    ],
  });
  const changed = structuredClone(original);
  const paragraph = changed.package.document.content.at(0);
  const table = changed.package.document.content.at(1);
  if (paragraph?.type !== "paragraph" || table?.type !== "table")
    throw new Error("Expected fixture blocks");
  const insertion = paragraph.content.at(0);
  if (insertion?.type !== "insertion") throw new Error("Expected insertion");
  const run = insertion.content.at(0);
  if (run?.type !== "run") throw new Error("Expected run");
  const row = table.rows.at(0);
  if (!row) throw new Error("Expected row");
  paragraph.formatting = { runProperties: {} };
  run.formatting = {};
  row.formatting = { sourceXml: "<w:trPr/>" };
  table.formatting = {
    sourceXml: "<w:tblPr/>",
    gridSourceXml: "<w:tblGrid><w:gridCol/></w:tblGrid>",
  };
  expect(reviewDifferences(original, changed)).toEqual({ messages: [], omitted: 0 });
  table.formatting.bidi = false;
  expect(reviewDifferences(original, changed).messages).not.toEqual([]);
  delete table.formatting.bidi;
  run.formatting = { bold: false };
  expect(reviewDifferences(original, changed).messages).not.toEqual([]);
});

test("derived current formatting agrees while an incorrect capture differs", () => {
  const original = fixture();
  const paragraph = original.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") throw new Error("Expected paragraph");
  const insertion = paragraph.content.at(0);
  if (insertion?.type !== "insertion") throw new Error("Expected insertion");
  const run = insertion.content.at(0);
  if (run?.type !== "run") throw new Error("Expected run");
  run.formatting = { bold: true };
  const change = { type: "runPropertyChange", info: { ...insertion.info, id: 42 } } as const;
  run.propertyChanges = [change];
  const captured = structuredClone(original);
  const capturedParagraph = captured.package.document.content.at(0);
  if (capturedParagraph?.type !== "paragraph") throw new Error("Expected paragraph");
  const capturedInsertion = capturedParagraph.content.at(0);
  if (capturedInsertion?.type !== "insertion") throw new Error("Expected insertion");
  const capturedRun = capturedInsertion.content.at(0);
  if (capturedRun?.type !== "run") throw new Error("Expected run");
  capturedRun.propertyChanges = [
    { ...change, previousFormatting: {}, currentFormatting: { bold: true } },
  ];
  expect(reviewDifferences(original, captured)).toEqual({ messages: [], omitted: 0 });
  capturedRun.propertyChanges = [{ ...change, currentFormatting: { bold: false } }];
  expect(reviewDifferences(original, captured).messages).not.toEqual([]);
});
