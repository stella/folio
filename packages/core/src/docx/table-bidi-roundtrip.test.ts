import { describe, expect, test } from "bun:test";

import { serializeTableFormatting } from "./serializer/tableSerializer";
import { parseTableProperties } from "./tableParser";
import type { XmlElement } from "./xmlParser";
import { parseXmlDocument } from "./xmlParser";

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const parseBidi = (xml: string) => {
  const tblPr = parseXmlDocument(`<w:tblPr ${W_NS}>${xml}</w:tblPr>`) as XmlElement | null;
  return parseTableProperties(tblPr)?.bidi;
};

describe("table bidiVisual tri-state", () => {
  test("preserves explicit RTL", () => {
    expect(parseBidi("<w:bidiVisual/>")).toBe(true);
    expect(serializeTableFormatting({ bidi: true })).toContain("<w:bidiVisual/>");
  });

  test("preserves explicit LTR so it can override an RTL table style", () => {
    expect(parseBidi('<w:bidiVisual w:val="0"/>')).toBe(false);
    expect(serializeTableFormatting({ bidi: false })).toContain('<w:bidiVisual w:val="0"/>');
  });

  test("keeps absence distinct from explicit LTR", () => {
    expect(parseBidi("")).toBeUndefined();
    expect(serializeTableFormatting({})).not.toContain("bidiVisual");
  });
});
