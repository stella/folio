/**
 * Every `ST_TextDirection` member survives a parse, a save and the editor
 * projection.
 *
 * `TABLE_CELL_TEXT_DIRECTION_VALUES` used to spell nine of the twelve. A cell
 * written `<w:textDirection w:val="tbLrV"/>` lost the attribute at parse time,
 * so it read as the table's own flow and saved without it: the column that had
 * been rotated came back horizontal.
 *
 * The sweep is over the generated list, so a schema refresh widens it.
 */

import { describe, expect, test } from "bun:test";

import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import type { Document, TableCellTextDirection } from "../types/document";
import { TABLE_CELL_TEXT_DIRECTION_VALUES } from "../types/documentEnumValues";

import { parseTableCellProperties } from "./tableParser";
import { parseXmlDocument } from "./xmlParser";

const WORD_NAMESPACE = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const parseCellTier = (direction: string): TableCellTextDirection | undefined => {
  const tcPr = parseXmlDocument(
    `<w:tcPr ${WORD_NAMESPACE}><w:textDirection w:val="${direction}"/></w:tcPr>`,
  );
  if (!tcPr) {
    throw new Error("fixture did not parse");
  }
  return parseTableCellProperties(tcPr)?.textDirection;
};

const documentWithDirection = (textDirection: TableCellTextDirection): Document => ({
  package: {
    document: {
      content: [
        {
          type: "table",
          rows: [
            {
              cells: [
                {
                  formatting: { textDirection },
                  content: [
                    {
                      type: "paragraph",
                      content: [{ type: "run", content: [{ type: "text", text: "Rotated" }] }],
                    },
                  ],
                },
              ],
            },
          ],
        },
      ],
    },
  },
});

const firstCellDirection = (document: Document): TableCellTextDirection | undefined => {
  const block = document.package.document.content.at(0);
  if (block?.type !== "table") {
    throw new Error("expected a table");
  }
  return block.rows.at(0)?.cells.at(0)?.formatting?.textDirection;
};

describe("ST_TextDirection", () => {
  // The sweeps below run over the model's own list, so they shrink with it.
  // `scripts/narrowed-enum-schema-types.test.ts` is what holds that list to the
  // enumeration; this names the three members it used to be missing.
  test("lrTb, lrTbV and tbLrV are members", () => {
    expect(TABLE_CELL_TEXT_DIRECTION_VALUES).toContain("lrTb");
    expect(TABLE_CELL_TEXT_DIRECTION_VALUES).toContain("lrTbV");
    expect(TABLE_CELL_TEXT_DIRECTION_VALUES).toContain("tbLrV");
  });

  test.each(TABLE_CELL_TEXT_DIRECTION_VALUES)("a cell's w:textDirection reads %s", (direction) => {
    expect(parseCellTier(direction)).toBe(direction);
  });

  test.each(TABLE_CELL_TEXT_DIRECTION_VALUES)("%s survives the editor projection", (direction) => {
    const original = documentWithDirection(direction);
    expect(firstCellDirection(fromProseDoc(toProseDoc(original), original))).toBe(direction);
  });

  test("a token outside the enumeration does not become a flow", () => {
    expect(parseCellTier("sideways")).toBeUndefined();
  });
});
