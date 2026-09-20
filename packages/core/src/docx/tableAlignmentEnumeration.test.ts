/**
 * Every `ST_JcTable` member survives a parse and a save, and `start` and `end`
 * survive as themselves.
 *
 * `w:tblPr/w:jc` and `w:trPr/w:jc` had a reader of their own: a chain of
 * literal comparisons that accepted `left`, `center` and `right`, folded
 * `start` onto `left` at the table and refused it outright at the row. A table
 * written `<w:jc w:val="start"/>` saved as `left`, which is a different value
 * in a right-to-left table, and one written `end` saved with no `w:jc` at all.
 *
 * The sweep is over the generated list, so a schema refresh widens it.
 */

import { describe, expect, test } from "bun:test";

import { TABLE_ALIGNMENTS } from "@stll/docx-core/model";

import { TABLE_JUSTIFICATION_VALUES } from "../types/documentEnumValues";

import {
  serializeTableFormatting,
  serializeTableRowFormatting,
} from "./serializer/tableSerializer";
import { parseTableProperties, parseTableRowProperties } from "./tableParser";
import { parseXmlDocument } from "./xmlParser";

const WORD_NAMESPACE = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const element = (markup: string) => {
  const parsed = parseXmlDocument(markup);
  if (!parsed) {
    throw new Error("fixture did not parse");
  }
  return parsed;
};

const parseTableAlignment = (value: string) =>
  parseTableProperties(element(`<w:tblPr ${WORD_NAMESPACE}><w:jc w:val="${value}"/></w:tblPr>`))
    ?.justification;

const parseRowAlignment = (value: string) =>
  parseTableRowProperties(element(`<w:trPr ${WORD_NAMESPACE}><w:jc w:val="${value}"/></w:trPr>`))
    ?.justification;

describe("ST_JcTable", () => {
  // `scripts/narrowed-enum-schema-types.test.ts` is what holds the picklist to
  // the enumeration; this names the two members the reader used to be missing.
  test("start and end are members", () => {
    expect(TABLE_JUSTIFICATION_VALUES).toBe(TABLE_ALIGNMENTS);
    expect(TABLE_JUSTIFICATION_VALUES).toContain("start");
    expect(TABLE_JUSTIFICATION_VALUES).toContain("end");
  });

  test.each(TABLE_ALIGNMENTS)("a table's w:jc reads %s", (value) => {
    expect(parseTableAlignment(value)).toBe(value);
  });

  test.each(TABLE_ALIGNMENTS)("a row's w:jc reads %s", (value) => {
    expect(parseRowAlignment(value)).toBe(value);
  });

  test.each(TABLE_ALIGNMENTS)("a table writes %s back as authored", (value) => {
    expect(serializeTableFormatting({ justification: value })).toContain(
      `<w:jc w:val="${value}"/>`,
    );
  });

  test.each(TABLE_ALIGNMENTS)("a row writes %s back as authored", (value) => {
    expect(serializeTableRowFormatting({ justification: value })).toContain(
      `<w:jc w:val="${value}"/>`,
    );
  });

  test("start is not read as left", () => {
    expect(parseTableAlignment("start")).toBe("start");
    expect(parseRowAlignment("start")).toBe("start");
  });

  test("end is read rather than dropped", () => {
    expect(parseTableAlignment("end")).toBe("end");
    expect(parseRowAlignment("end")).toBe("end");
  });

  test("a token outside the enumeration is not read as a placement", () => {
    expect(parseTableAlignment("both")).toBeUndefined();
    expect(parseRowAlignment("both")).toBeUndefined();
  });
});
