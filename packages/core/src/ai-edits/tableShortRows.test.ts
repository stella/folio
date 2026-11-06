/**
 * A row may stop short of its table's width with no `w:gridAfter` to say so.
 * Word draws nothing past its last cell, and an edit must not change that.
 */

import { describe, expect, test } from "bun:test";
import { TableMap } from "prosemirror-tables";

import { ZipArchive } from "@stll/docx-core/zip";

import {
  createHarnessState,
  parseShapeDocument,
} from "../__tests__/editorHarness";
import { createDocx } from "../docx/rezip";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
const cell = (text: string) =>
  `<w:tc><w:tcPr/><w:p w14:paraId="1000000${text}"><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
const block = (text: string) => `1000000${text}`;

/** Row 1 fills the grid; rows 2 and 3 stop short of it. */
const shortRowsDocx = async (): Promise<ArrayBuffer> => {
  const zip = await ZipArchive.read(await createDocx(createEmptyDocument()));
  zip.set(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${NS}><w:body><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid><w:tr>${cell("1")}${cell("2")}${cell("3")}</w:tr><w:tr>${cell("4")}</w:tr><w:tr>${cell("5")}${cell("6")}</w:tr></w:tbl><w:p w14:paraId="10000007"><w:r><w:t>Tail</w:t></w:r></w:p></w:body></w:document>`,
  );
  return (await zip.generate()).slice().buffer;
};

/** The saved rows, each as the text of its cells. */
const savedRows = async (buffer: ArrayBuffer): Promise<string[][]> => {
  const xml =
    (await (
      await ZipArchive.read(buffer)
    )
      .entry("word/document.xml")
      ?.text()) ?? "";
  return [...xml.matchAll(/<w:tr[ >].*?<\/w:tr>/gsu)].map(([row]) =>
    [...row.matchAll(/<w:tc>.*?<\/w:tc>/gsu)].map(([saved]) =>
      [...saved.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/gu)]
        .map(([, text]) => text)
        .join(""),
    ),
  );
};

const open = async () => FolioDocxReviewer.fromBuffer(await shortRowsDocx());

const apply = (
  reviewer: FolioDocxReviewer,
  mode: "direct" | "tracked-changes",
  operation: Record<string, unknown>,
) =>
  reviewer.applyDocumentOperations(
    { version: 1, mode, operations: [{ id: "edit", ...operation }] } as never,
    { undefinedReferences: "refuse" },
  );

const MODES = ["direct", "tracked-changes"] as const;
const SOURCE = [["1", "2", "3"], ["4"], ["5", "6"]];

describe("a table with rows shorter than its width", () => {
  test("opens as a rectangle, so the table map has no empty slot", async () => {
    const state = createHarnessState(
      await parseShapeDocument(new Uint8Array(await shortRowsDocx())),
      "editing",
    );
    const map = TableMap.get(state.doc.child(0));
    expect(map.problems).toBeNull();
    expect(map.map).not.toContain(0);
  });

  test("lists no block for a position the row omits", async () => {
    const reviewer = await open();
    expect(reviewer.getContent().map(({ text }) => text)).toEqual([
      "1",
      "2",
      "3",
      "4",
      "5",
      "6",
      "Tail",
    ]);
  });

  test("saves untouched rows as it found them", async () => {
    const reviewer = await open();
    expect(await savedRows(await reviewer.toBuffer())).toEqual(SOURCE);
  });

  for (const mode of MODES) {
    test(`keeps its rows through a text edit (${mode})`, async () => {
      const reviewer = await open();
      const result = apply(reviewer, mode, {
        type: "insertAfterBlock",
        blockId: block("1"),
        text: "Added",
      });
      expect(result.applied).toHaveLength(1);
      const rows = await savedRows(await reviewer.toBuffer());
      // The added paragraph is in the first cell's row; nothing else moved.
      expect(rows.slice(1)).toEqual(SOURCE.slice(1));
      expect(rows[0]).toHaveLength(3);

      reviewer.rejectAll();
      expect(await savedRows(await reviewer.toBuffer())).toEqual(
        mode === "direct" ? rows : SOURCE,
      );
    });

    for (const anchor of ["1", "4", "5", "6"]) {
      for (const position of ["before", "after"] as const) {
        test(`takes a row ${position} cell ${anchor} (${mode})`, async () => {
          const reviewer = await open();
          const result = apply(reviewer, mode, {
            type: "insertTableRow",
            blockId: block(anchor),
            position,
            cellTexts: ["New"],
          });
          expect(result.applied).toHaveLength(1);
          const rows = await savedRows(await reviewer.toBuffer());
          expect(rows.filter((row) => row.includes("New"))).toEqual([
            ["New", "", ""],
          ]);
          expect(rows.filter((row) => !row.includes("New"))).toEqual(SOURCE);
        });
      }
    }
  }

  test("takes a column through a short row", async () => {
    const reviewer = await open();
    const result = apply(reviewer, "direct", {
      type: "insertTableColumn",
      blockId: block("1"),
      position: "after",
    });
    expect(result.applied).toHaveLength(1);
    expect(await savedRows(await reviewer.toBuffer())).toEqual([
      ["1", "", "2", "3"],
      ["4", ""],
      ["5", "", "6"],
    ]);
  });

  describe("losing the only cell a row has", () => {
    const EXPECTED = [["2", "3"], ["6"]];

    test("removes the row in a direct edit", async () => {
      const reviewer = await open();
      const result = apply(reviewer, "direct", {
        type: "deleteTableColumn",
        blockId: block("1"),
      });
      expect(result.applied).toHaveLength(1);
      expect(await savedRows(await reviewer.toBuffer())).toEqual(EXPECTED);
    });

    test("removes the row when the tracked deletion is accepted", async () => {
      const reviewer = await open();
      const result = apply(reviewer, "tracked-changes", {
        type: "deleteTableColumn",
        blockId: block("1"),
      });
      expect(result.applied).toHaveLength(1);
      reviewer.acceptAll();
      expect(await savedRows(await reviewer.toBuffer())).toEqual(EXPECTED);
    });

    test("restores it when the tracked deletion is rejected", async () => {
      const reviewer = await open();
      apply(reviewer, "tracked-changes", {
        type: "deleteTableColumn",
        blockId: block("1"),
      });
      reviewer.rejectAll();
      expect(await savedRows(await reviewer.toBuffer())).toEqual(SOURCE);
    });
  });
});
