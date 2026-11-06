/**
 * A bookmark between rows or between cells keeps its place among them, and a
 * document holding one saves after any of them comes or goes.
 */

import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { createDocx } from "../docx/rezip";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
const cell = (text: string) =>
  `<w:tc><w:tcPr/><w:p w14:paraId="1000000${text}"><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
const row = (...cells: string[]) => `<w:tr>${cells.map(cell).join("")}</w:tr>`;
const mark = (id: number) =>
  `<w:bookmarkStart w:id="${id}" w:name="b${id}"/><w:bookmarkEnd w:id="${id}"/>`;

const open = async (rows: string): Promise<FolioDocxReviewer> => {
  const zip = await ZipArchive.read(await createDocx(createEmptyDocument()));
  zip.set(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${NS}><w:body><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>${rows}</w:tbl><w:p w14:paraId="10000009"><w:r><w:t>Tail</w:t></w:r></w:p></w:body></w:document>`,
  );
  return FolioDocxReviewer.fromBuffer((await zip.generate()).slice().buffer);
};

const savedTable = async (reviewer: FolioDocxReviewer) => {
  const xml =
    (await (
      await ZipArchive.read(await reviewer.toBuffer())
    )
      .entry("word/document.xml")
      ?.text()) ?? "";
  return [
    ...xml.matchAll(
      /<w:bookmarkStart[^>]*w:name="(b\d)"|<w:t(?: [^>]*)?>(\d)</gu,
    ),
  ].map(([, name, text]) => name ?? text);
};

const apply = (
  reviewer: FolioDocxReviewer,
  mode: "direct" | "tracked-changes",
  operation: Record<string, unknown>,
) => {
  const result = reviewer.applyDocumentOperations(
    { version: 1, mode, operations: [{ id: "edit", ...operation }] } as never,
    { undefinedReferences: "refuse" },
  );
  expect(result.applied).toHaveLength(1);
};

const BETWEEN_ROWS = `${row("1", "2")}${mark(1)}${row("3", "4")}${row("5", "6")}${mark(2)}`;
const BETWEEN_CELLS = `<w:tr>${cell("1")}${mark(1)}${cell("2")}${mark(2)}</w:tr>${row("3", "4")}`;

const REMOVE_ROW = { type: "deleteTableRow", blockId: "10000005" };
const REMOVE_COLUMN = { type: "deleteTableColumn", blockId: "10000002" };

describe("a bookmark between a table's rows", () => {
  test("stays before the row that followed it when an earlier row goes", async () => {
    const reviewer = await open(BETWEEN_ROWS);
    apply(reviewer, "direct", { type: "deleteTableRow", blockId: "10000001" });
    expect(await savedTable(reviewer)).toEqual([
      "b1",
      "3",
      "4",
      "5",
      "6",
      "b2",
    ]);
  });

  test("survives losing the row after it", async () => {
    const reviewer = await open(BETWEEN_ROWS);
    apply(reviewer, "direct", REMOVE_ROW);
    expect(await savedTable(reviewer)).toEqual([
      "1",
      "2",
      "b1",
      "3",
      "4",
      "b2",
    ]);
  });

  test("survives an accepted row deletion, and a rejected one restores it", async () => {
    const accepted = await open(BETWEEN_ROWS);
    apply(accepted, "tracked-changes", REMOVE_ROW);
    accepted.acceptAll();
    expect(await savedTable(accepted)).toEqual([
      "1",
      "2",
      "b1",
      "3",
      "4",
      "b2",
    ]);

    const rejected = await open(BETWEEN_ROWS);
    apply(rejected, "tracked-changes", REMOVE_ROW);
    rejected.rejectAll();
    expect(await savedTable(rejected)).toEqual([
      "1",
      "2",
      "b1",
      "3",
      "4",
      "5",
      "6",
      "b2",
    ]);
  });

  test("keeps its neighbours when a row is added elsewhere", async () => {
    const reviewer = await open(BETWEEN_ROWS);
    apply(reviewer, "direct", {
      type: "insertTableRow",
      blockId: "10000001",
      position: "before",
      cellTexts: ["7", "8"],
    });
    expect(await savedTable(reviewer)).toEqual([
      "7",
      "8",
      "1",
      "2",
      "b1",
      "3",
      "4",
      "5",
      "6",
      "b2",
    ]);
  });
});

describe("a bookmark between a row's cells", () => {
  test("survives losing the cell after it", async () => {
    const reviewer = await open(BETWEEN_CELLS);
    apply(reviewer, "direct", REMOVE_COLUMN);
    expect(await savedTable(reviewer)).toEqual(["1", "b1", "b2", "3"]);
  });

  test("survives an accepted column deletion", async () => {
    const reviewer = await open(BETWEEN_CELLS);
    apply(reviewer, "tracked-changes", REMOVE_COLUMN);
    reviewer.acceptAll();
    expect(await savedTable(reviewer)).toEqual(["1", "b1", "b2", "3"]);
  });
});

describe("a bookmark that opens in a cell and closes on the table or a row", () => {
  const endsOf = async (reviewer: FolioDocxReviewer) => {
    const xml =
      (await (
        await ZipArchive.read(await reviewer.toBuffer())
      )
        .entry("word/document.xml")
        ?.text()) ?? "";
    return [...xml.matchAll(/<w:bookmarkEnd w:id="7"\/>(.{0,7})/gu)].map(
      ([, next]) => next,
    );
  };

  const OPENING_CELL = `<w:tc><w:tcPr/><w:p w14:paraId="10000001"><w:bookmarkStart w:id="7" w:name="b7"/><w:r><w:t>1</w:t></w:r></w:p></w:tc>`;
  test.each([
    [
      "on the table",
      `<w:tr>${OPENING_CELL}${cell("2")}</w:tr>${row("3", "4")}<w:bookmarkEnd w:id="7"/>`,
    ],
    [
      "on its own row",
      `<w:tr>${OPENING_CELL}<w:bookmarkEnd w:id="7"/>${cell("2")}</w:tr>${row("3", "4")}`,
    ],
  ])("closing %s, an edit keeps its one end", async (_where, table) => {
    const reviewer = await open(table);
    apply(reviewer, "direct", {
      type: "insertTableRow",
      blockId: "10000001",
      position: "after",
      cellTexts: ["x", "y"],
    });
    expect(await endsOf(reviewer)).toHaveLength(1);
  });
});
