import { describe, expect, test } from "bun:test";

import { parseRunProperties } from "../runParser";
import { serializeTextFormatting } from "../serializer/runSerializer";
import { parseXml } from "../xmlParser";

const parseScale = (value: string) => {
  const parsed = parseXml(
    `<w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:w w:val="${value}"/></w:rPr>`,
  );
  const runProperties = parsed.elements.at(0);
  if (!runProperties || runProperties.type !== "element") {
    throw new Error("Expected run properties element");
  }
  return parseRunProperties(runProperties, null);
};

describe("horizontal text scale round-trip", () => {
  test.each([
    ["0", 0],
    ["0%", 0],
    ["50%", 50],
    ["600", 600],
    ["600%", 600],
    [" 50% ", 50],
  ])("preserves schema-valid scale %s", (value, expected) => {
    const formatting = parseScale(value);

    expect(formatting?.scale).toBe(expected);
    expect(serializeTextFormatting(formatting)).toContain(`<w:w w:val="${expected}"/>`);
  });

  test.each([
    "",
    "   ",
    "-1",
    "601",
    "601%",
    "0garbage",
    "1e2",
    "600oops",
    "0x10",
    "Infinity",
    "not-a-scale",
  ])("drops malformed scale %s at parse and serialization boundaries", (value) => {
    const formatting = parseScale(value);

    expect(formatting?.scale).toBeUndefined();
    expect(serializeTextFormatting(formatting)).not.toContain("<w:w ");
  });
});
