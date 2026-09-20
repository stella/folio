import { describe, expect, test } from "bun:test";

import { parseRunProperties } from "../runParser";
import { serializeTextFormatting } from "../serializer/textFormattingSerializer";
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
  ])("takes no scale from malformed %s, and keeps its bytes", (value) => {
    const formatting = parseScale(value);

    // The reader admits no scale, so nothing is rebuilt from the model — but
    // the element the author wrote is not folio's to discard, so it comes back
    // verbatim from the sink rather than being dropped.
    expect(formatting?.scale).toBeUndefined();
    expect(serializeTextFormatting(formatting)).toContain(`<w:w w:val="${value}"/>`);
  });
});
