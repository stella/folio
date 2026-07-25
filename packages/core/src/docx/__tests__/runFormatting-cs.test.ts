import { describe, expect, test } from "bun:test";

import { parseRunProperties } from "../runParser";
import { serializeTextFormatting } from "../serializer/runSerializer";
import { parseXml } from "../xmlParser";

const parseFormatting = (innerXml: string) => {
  const parsed = parseXml(
    `<w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${innerXml}</w:rPr>`,
  );
  const runProperties = parsed.elements.at(0);
  if (!runProperties || runProperties.type !== "element") {
    throw new TypeError("Expected run properties element");
  }
  return parseRunProperties(runProperties, null);
};

describe("complex-script run formatting round-trip", () => {
  test("preserves an explicit false override", () => {
    const formatting = parseFormatting('<w:cs w:val="0"/>');

    expect(formatting?.cs).toBe(false);
    expect(serializeTextFormatting(formatting)).toContain('<w:cs w:val="0"/>');
  });

  test("preserves an enabled override", () => {
    const formatting = parseFormatting("<w:cs/>");

    expect(formatting?.cs).toBe(true);
    expect(serializeTextFormatting(formatting)).toContain("<w:cs/>");
  });

  test("does not materialize an absent override", () => {
    const formatting = parseFormatting("<w:b/>");

    expect(formatting?.cs).toBeUndefined();
    expect(serializeTextFormatting(formatting)).not.toContain("<w:cs");
  });
});
