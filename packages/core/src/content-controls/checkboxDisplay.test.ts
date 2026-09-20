import { describe, expect, test } from "bun:test";

import { checkboxDisplayContent } from "./checkboxDisplay";

const propertiesXml = (states: string): string =>
  `<w:sdtPr><w14:checkbox>${states}</w14:checkbox></w:sdtPr>`;

describe("checkboxDisplayContent", () => {
  test("uses the state matching the requested value", () => {
    const raw = propertiesXml(
      '<w14:checkedState w14:val="F0FE" w14:font="Wingdings"/><w14:uncheckedState w14:val="F0A8" w14:font="Wingdings 2"/>',
    );
    expect(checkboxDisplayContent(raw, true)).toEqual({
      type: "symbol",
      char: "F0FE",
      font: "Wingdings",
    });
    expect(checkboxDisplayContent(raw, false)).toEqual({
      type: "symbol",
      char: "F0A8",
      font: "Wingdings 2",
    });
  });

  test("normalizes short BMP codes and preserves supplementary Unicode as text", () => {
    expect(checkboxDisplayContent(propertiesXml('<w14:checkedState w14:val="41"/>'), true)).toEqual(
      { type: "symbol", char: "0041", font: "MS Gothic" },
    );
    expect(
      checkboxDisplayContent(propertiesXml('<w14:checkedState w14:val="1F5F9"/>'), true),
    ).toEqual({ type: "text", text: "🗹" });
  });

  test("falls back for absent, malformed, or foreign-namespace state declarations", () => {
    const expected = { type: "symbol", char: "2612", font: "MS Gothic" };
    expect(checkboxDisplayContent(undefined, true)).toEqual(expected);
    expect(
      checkboxDisplayContent(propertiesXml('<w14:checkedState w14:val="not-hex"/>'), true),
    ).toEqual(expected);
    expect(
      checkboxDisplayContent(
        '<w:sdtPr xmlns:x="urn:foreign"><w14:checkbox><x:checkedState x:val="F0FE" x:font="Wingdings"/></w14:checkbox></w:sdtPr>',
        true,
      ),
    ).toEqual(expected);
    expect(
      checkboxDisplayContent(
        `<w:sdtPr>${" ".repeat(65_536)}<w14:checkbox><w14:checkedState w14:val="F0FE"/></w14:checkbox></w:sdtPr>`,
        true,
      ),
    ).toEqual(expected);
  });
});
