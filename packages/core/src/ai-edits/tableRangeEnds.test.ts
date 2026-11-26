/**
 * Removing a table row or a whole table takes its content, not the end of a
 * bookmark or move range reaching past it: that end stays where the row or
 * table was, so the document still saves with every range whole.
 */

import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { createDocx } from "../docx/rezip";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
const BY = 'w:author="A" w:date="2026-01-01T00:00:00Z"';
const START = `<w:bookmarkStart w:id="4" w:name="b"/><w:moveToRangeStart w:id="5" ${BY} w:name="m"/>`;
const END = '<w:bookmarkEnd w:id="4"/><w:moveToRangeEnd w:id="5"/>';
const MOVED_FROM = `<w:p w14:paraId="30000001"><w:moveFromRangeStart w:id="1" ${BY} w:name="m"/><w:moveFrom w:id="2" ${BY}><w:r><w:t>Moved</w:t></w:r></w:moveFrom><w:moveFromRangeEnd w:id="1"/></w:p>`;

const paragraph = (id: string, inner = "") =>
  `<w:p w14:paraId="${id}">${inner}<w:r><w:t>P${id.slice(-1)}</w:t></w:r></w:p>`;
const cell = (id: string, inner = "") =>
  `<w:tc><w:tcPr/>${paragraph(id, inner)}</w:tc>`;
const table = (rows: readonly (readonly [string, string])[]) =>
  `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>${rows
    .map(([id, inner]) => `<w:tr>${cell(id, inner)}</w:tr>`)
    .join("")}</w:tbl>`;

const open = async (body: string): Promise<FolioDocxReviewer> => {
  const zip = await ZipArchive.read(await createDocx(createEmptyDocument()));
  zip.set(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${NS}><w:body>${MOVED_FROM}${body}</w:body></w:document>`,
  );
  return FolioDocxReviewer.fromBuffer((await zip.generate()).slice().buffer, {
    session: "canonical",
  });
};

const apply = (
  reviewer: FolioDocxReviewer,
  operation: Record<string, unknown>,
) =>
  expect(
    reviewer.applyDocumentOperations(
      { version: 1, mode: "direct", operations: [{ id: "e", ...operation }] },
      { undefinedReferences: "refuse" },
    ).applied,
  ).toHaveLength(1);

/** The range boundaries the saved body holds, in document order. */
const savedBoundaries = async (reviewer: FolioDocxReviewer) => {
  const xml =
    (await (
      await ZipArchive.read(await reviewer.toBuffer())
    )
      .entry("word/document.xml")
      ?.text()) ?? "";
  return [
    ...xml.matchAll(/<w:(bookmark|moveToRange)(Start|End) w:id="(\d)"/gu),
  ].map(([, family, side]) => `${family}${side}`);
};

const WHOLE = [
  "bookmarkStart",
  "moveToRangeStart",
  "bookmarkEnd",
  "moveToRangeEnd",
];

describe("a range reaching past a removed table row", () => {
  test.each([
    [
      "the first row",
      "10000001",
      [
        ["10000001", START],
        ["10000002", ""],
      ],
    ],
    [
      "the last row",
      "10000002",
      [
        ["10000001", ""],
        ["10000002", START],
      ],
    ],
    ["the only row", "10000001", [["10000001", START]]],
  ] as const)("keeps its start when %s goes", async (_name, blockId, rows) => {
    const reviewer = await open(
      `${paragraph("20000001")}${table(rows)}${paragraph("20000002", END)}`,
    );
    apply(reviewer, { type: "deleteTableRow", blockId });
    expect(await savedBoundaries(reviewer)).toEqual(WHOLE);
  });

  test("keeps its end when the row holding it goes", async () => {
    const reviewer = await open(
      `${paragraph("20000001", START)}${table([
        ["10000001", END],
        ["10000002", ""],
      ])}${paragraph("20000002")}`,
    );
    apply(reviewer, { type: "deleteTableRow", blockId: "10000001" });
    expect(await savedBoundaries(reviewer)).toEqual(WHOLE);
  });

  test("keeps its start when the whole table goes", async () => {
    const reviewer = await open(
      `${table([["10000001", START]])}${paragraph("20000002", END)}`,
    );
    apply(reviewer, { type: "deleteTable", blockId: "10000001" });
    expect(await savedBoundaries(reviewer)).toEqual(WHOLE);
  });

  // The shrunk consumer flow: a table inserted after a moved paragraph,
  // the move accepted, the paragraph mark's suggestion rejected (its words
  // run into the table's first cell), then that row removed.
  test("saves after the moved paragraph's row goes", async () => {
    const reviewer = await open(
      `<w:p w14:paraId="20000001"><w:pPr><w:rPr><w:moveTo w:id="3" ${BY}/></w:rPr></w:pPr>${START}<w:moveTo w:id="6" ${BY}><w:r><w:t>Moved</w:t></w:r></w:moveTo></w:p>${END}${paragraph("20000002")}`,
    );
    apply(reviewer, {
      type: "insertTable",
      blockId: "20000001",
      rows: [
        ["Term", "Value"],
        ["Period", "12 months"],
      ],
    });
    const moved = reviewer
      .getChanges()
      .find((change) => change.type === "insertion");
    expect(moved && reviewer.acceptChange(moved)).toBe(true);
    reviewer.rejectAll();
    const first = reviewer
      .getContent()
      .find((block) => "text" in block && block.text.includes("Term"));
    apply(reviewer, { type: "deleteTableRow", blockId: first?.id });
    expect(await savedBoundaries(reviewer)).toEqual(WHOLE);
  });
});
