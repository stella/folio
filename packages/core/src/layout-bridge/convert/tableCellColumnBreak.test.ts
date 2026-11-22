import { expect, test } from "bun:test";

import { parseDocumentBody } from "../../docx/documentParser";
import {
  fixedCharWidth,
  withFakeTextMeasure,
} from "../../layout-engine/measure/__tests__/fakeTextMeasure";
import { measureTableBlock } from "../../layout-engine/measure/measureBlocks";
import { toProseDoc } from "../../prosemirror/conversion/toProseDoc";
import { toFlowBlocks } from "./toFlowBlocks";

test.each([
  ["column", 1],
  ["textWrapping", 2],
] as const)(
  "a %s break in a table cell yields %i visual lines",
  (type, count) => {
    const document = parseDocumentBody(
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' +
        "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Alpha</w:t>" +
        `<w:br w:type="${type}"/><w:t>Beta</w:t></w:r></w:p></w:tc></w:tr></w:tbl>` +
        "</w:body></w:document>",
    );
    const table = toFlowBlocks(toProseDoc({ package: { document } })).find(
      (block) => block.kind === "table",
    );
    if (!table) {
      throw new Error("Expected a table in the document body");
    }
    withFakeTextMeasure(
      () => {
        const measure = measureTableBlock(table, 500);
        const paragraph = measure.rows.at(0)?.cells.at(0)?.blocks.at(0);
        expect(
          paragraph?.kind === "paragraph" ? paragraph.lines : [],
        ).toHaveLength(count);
      },
      { charWidth: fixedCharWidth(8) },
    );
  },
);
