import { expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { parseDocx } from "../../docx/parser";
import { createEmptyDocx } from "../../docx/rezip";
import {
  fixedCharWidth,
  withFakeTextMeasure,
} from "../../layout-engine/measure/__tests__/fakeTextMeasure";
import { measureTableBlock } from "../../layout-engine/measure/measureBlocks";
import { toProseDoc } from "../../prosemirror/conversion/toProseDoc";
import { toFlowBlocks } from "./toFlowBlocks";

const LABEL = "Blue green orange";
const CHAR_WIDTH_PX = 8;

type TableCase = {
  layout: "autofit" | "fixed";
  firstWidth: "auto" | "dxa";
  noWrap: boolean;
  columnWidthTwips?: number;
  keepSecondWidth?: boolean;
};

const tableFixture = async ({
  layout,
  firstWidth,
  noWrap,
  columnWidthTwips = 900,
  keepSecondWidth = true,
}: TableCase) => {
  const cell = (text: string, properties: string) =>
    `<w:tc><w:tcPr>${properties}</w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
  const table =
    `<w:tbl><w:tblPr><w:tblW w:w="${columnWidthTwips * 2}" w:type="dxa"/>` +
    `<w:tblLayout w:type="${layout}"/>` +
    '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/>' +
    '<w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>' +
    `</w:tblPr><w:tblGrid><w:gridCol w:w="${columnWidthTwips}"/><w:gridCol w:w="${columnWidthTwips}"/></w:tblGrid><w:tr>` +
    cell(
      LABEL,
      `<w:tcW w:w="${firstWidth === "auto" ? 0 : columnWidthTwips}" w:type="${firstWidth}"/>` +
        (noWrap ? "<w:noWrap/>" : ""),
    ) +
    cell(
      "1",
      `<w:tcW w:w="${columnWidthTwips}" w:type="dxa"/>` +
        (keepSecondWidth ? "<w:noWrap/>" : ""),
    ) +
    "</w:tr></w:tbl>";
  const archive = await ZipArchive.read(await createEmptyDocx());
  archive.set(
    "word/document.xml",
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      `<w:body>${table}<w:sectPr/></w:body></w:document>`,
  );
  const { buffer } = await archive.generate({ compression: "store" });
  const document = await parseDocx(buffer, { preloadFonts: false });
  const block = toFlowBlocks(toProseDoc(document)).find(
    (candidate) => candidate.kind === "table",
  );
  if (!block) {
    throw new Error("Expected a table in the fixture");
  }
  return block;
};

test("an AutoFit cell with an automatic width expands to keep its text together", async () => {
  const table = await tableFixture({
    layout: "autofit",
    firstWidth: "auto",
    noWrap: true,
  });
  withFakeTextMeasure(
    () => {
      const measure = measureTableBlock(table, 500);
      const paragraph = measure.rows.at(0)?.cells.at(0)?.blocks.at(0);
      expect(
        paragraph?.kind === "paragraph" ? paragraph.lines : [],
      ).toHaveLength(1);
      expect(measure.columnWidths.at(0)).toBeGreaterThanOrEqual(
        LABEL.length * CHAR_WIDTH_PX,
      );
      expect(measure.columnWidths.at(1)).toBeCloseTo(60);
      expect(measure.totalWidth).toBeGreaterThan(120);
    },
    { charWidth: fixedCharWidth(CHAR_WIDTH_PX) },
  );
});

test("an AutoFit cell can borrow unused room from its neighbor", async () => {
  const table = await tableFixture({
    layout: "autofit",
    firstWidth: "auto",
    noWrap: true,
    columnWidthTwips: 1350,
    keepSecondWidth: false,
  });
  withFakeTextMeasure(
    () => {
      const measure = measureTableBlock(table, 500);
      const unwrappedWidth = LABEL.length * CHAR_WIDTH_PX;
      expect(measure.columnWidths).toEqual([
        unwrappedWidth,
        180 - unwrappedWidth,
      ]);
      expect(measure.totalWidth).toBeCloseTo(180);
    },
    { charWidth: fixedCharWidth(CHAR_WIDTH_PX) },
  );
});

test.each([
  { layout: "autofit", firstWidth: "dxa", noWrap: true },
  { layout: "fixed", firstWidth: "auto", noWrap: true },
  { layout: "autofit", firstWidth: "auto", noWrap: false },
] as const)("keeps the recorded grid and wraps with %j", async (options) => {
  const table = await tableFixture(options);
  withFakeTextMeasure(
    () => {
      const measure = measureTableBlock(table, 500);
      const paragraph = measure.rows.at(0)?.cells.at(0)?.blocks.at(0);
      expect(
        paragraph?.kind === "paragraph" ? paragraph.lines.length : 0,
      ).toBeGreaterThan(1);
      expect(measure.columnWidths).toEqual([60, 60]);
      expect(measure.totalWidth).toBeCloseTo(120);
    },
    { charWidth: fixedCharWidth(CHAR_WIDTH_PX) },
  );
});
