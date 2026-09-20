/**
 * The two spellings of one `ST_TextDirection` flow reach the same rendering.
 *
 * The enumeration carries twelve tokens for six flows: Strict spells them
 * `tb`, `rl`, `lr`, `tbV`, `rlV` and `lrV`, and Transitional adds `lrTb`,
 * `tbRl`, `btLr`, `lrTbV`, `tbRlV` and `tbLrV` for the same six (ECMA-376
 * Part 4 §14.11.7, Part 1 §17.18.93). Folio paired them by the letters in the
 * token instead: `tb` read as a vertical flow because the long vertical
 * spellings start `tb`, and `rl` read as right-to-left inline text because
 * `tbRl` ends `Rl`. Every one of the six Strict spellings rendered as
 * something other than its Transitional twin, so a Strict-conformance package
 * painted its rotated cells flat and its flat cells rotated.
 *
 * The sweep is over the generated flow map, so a schema refresh widens it.
 */

import { describe, expect, test } from "bun:test";

import {
  TEXT_DIRECTION_FLOW_BY_TOKEN,
  TEXT_DIRECTION_FLOWS,
  TEXT_DIRECTIONS,
  type TextDirection,
  type TextDirectionFlow,
} from "@stll/docx-core/model";

import { cellTextRotationDegrees } from "../layout-painter/renderTable";
import { tableCellToStyle } from "../utils/formatToStyle";
import { textFlowCss } from "../utils/textDirectionFlow";

import { parseSectionProperties } from "./sectionParser";
import { serializeTableCellFormatting } from "./serializer/tableSerializer";
import { parseTableCellProperties } from "./tableParser";
import { parseXmlDocument } from "./xmlParser";

const WORD_NAMESPACE = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

/**
 * The Transitional spelling of each flow, from Part 4 §14.11.7. Written out
 * rather than derived: this is the table the fix is about, and deriving it
 * from the map under test would assert nothing.
 */
const TRANSITIONAL_SPELLING = {
  tb: "lrTb",
  rl: "tbRl",
  lr: "btLr",
  tbV: "lrTbV",
  rlV: "tbRlV",
  lrV: "tbLrV",
} as const satisfies Record<TextDirectionFlow, TextDirection>;

const parseCellDirection = (direction: string): TextDirection | undefined => {
  const tcPr = parseXmlDocument(
    `<w:tcPr ${WORD_NAMESPACE}><w:textDirection w:val="${direction}"/></w:tcPr>`,
  );
  if (!tcPr) {
    throw new Error("fixture did not parse");
  }
  return parseTableCellProperties(tcPr)?.textDirection;
};

const parseSectionDirection = (direction: string): TextDirection | undefined => {
  const sectPr = parseXmlDocument(
    `<w:sectPr ${WORD_NAMESPACE}><w:textDirection w:val="${direction}"/></w:sectPr>`,
  );
  if (!sectPr) {
    throw new Error("fixture did not parse");
  }
  return parseSectionProperties(sectPr).textDirection;
};

const flowPairs = TEXT_DIRECTION_FLOWS.map(
  (flow) =>
    [flow, TRANSITIONAL_SPELLING[flow]] as const satisfies readonly [
      TextDirectionFlow,
      TextDirection,
    ],
);

describe("ST_TextDirection spellings", () => {
  test("the enumeration is twelve tokens over six flows", () => {
    expect(TEXT_DIRECTIONS).toHaveLength(12);
    expect(TEXT_DIRECTION_FLOWS).toHaveLength(6);
    expect(new Set(Object.values(TEXT_DIRECTION_FLOW_BY_TOKEN)).size).toBe(6);
  });

  test.each(flowPairs)("%s and %s are one flow", (strict, transitional) => {
    expect(TEXT_DIRECTION_FLOW_BY_TOKEN[strict]).toBe(strict);
    expect(TEXT_DIRECTION_FLOW_BY_TOKEN[transitional]).toBe(strict);
  });

  test.each(flowPairs)("%s and %s turn the same quarter turn", (strict, transitional) => {
    expect(cellTextRotationDegrees(strict)).toBe(cellTextRotationDegrees(transitional));
  });

  test.each(flowPairs)("%s and %s render the same CSS flow", (strict, transitional) => {
    expect(textFlowCss(strict)).toEqual(textFlowCss(transitional));
  });

  test.each(flowPairs)("%s and %s style a cell the same", (strict, transitional) => {
    expect(tableCellToStyle({ textDirection: strict })).toEqual(
      tableCellToStyle({ textDirection: transitional }),
    );
  });

  test.each(TEXT_DIRECTIONS)("a cell's w:textDirection reads %s", (direction) => {
    expect(parseCellDirection(direction)).toBe(direction);
  });

  test.each(TEXT_DIRECTIONS)("a section's w:textDirection reads %s", (direction) => {
    expect(parseSectionDirection(direction)).toBe(direction);
  });

  test.each(TEXT_DIRECTIONS)("%s is written back as authored", (direction) => {
    expect(serializeTableCellFormatting({ textDirection: direction })).toContain(
      `<w:textDirection w:val="${direction}"/>`,
    );
  });

  // The six the pairing used to get wrong, named one at a time. Each one is a
  // Strict spelling that rendered as its letters rather than as its flow.
  test("tb is the horizontal flow, not a vertical one", () => {
    expect(cellTextRotationDegrees("tb")).toBe(0);
    expect(textFlowCss("tb").writingMode).toBe("horizontal-tb");
  });

  test("tbV is the horizontal flow with East Asian characters turned", () => {
    expect(cellTextRotationDegrees("tbV")).toBe(0);
    expect(textFlowCss("tbV").writingMode).toBe("horizontal-tb");
  });

  test("rl is a vertical flow, not right-to-left inline text", () => {
    expect(cellTextRotationDegrees("rl")).toBe(90);
    expect(textFlowCss("rl").writingMode).toBe("vertical-rl");
    expect(tableCellToStyle({ textDirection: "rl" }).direction).toBeUndefined();
  });

  test("rlV is a vertical flow, not right-to-left inline text", () => {
    expect(cellTextRotationDegrees("rlV")).toBe(90);
    expect(textFlowCss("rlV").writingMode).toBe("vertical-rl");
    expect(tableCellToStyle({ textDirection: "rlV" }).direction).toBeUndefined();
  });

  test("lr runs bottom to top, so it turns counter-clockwise", () => {
    expect(cellTextRotationDegrees("lr")).toBe(-90);
    expect(textFlowCss("lr")).toEqual({ writingMode: "vertical-lr", rotateDegrees: 180 });
  });

  test("lrV is a vertical flow, not a horizontal one", () => {
    expect(cellTextRotationDegrees("lrV")).toBe(90);
    expect(textFlowCss("lrV").writingMode).toBe("vertical-lr");
  });
});
