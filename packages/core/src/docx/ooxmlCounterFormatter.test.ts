import { describe, expect, test } from "bun:test";

import { NUMBER_FORMAT_VALUES } from "../types/documentEnumValues";
import { parseBlockContent } from "./blockContentParser";
import { parseNumbering } from "./numberingParser";
import { formatOoxmlCounter } from "./ooxmlCounterFormatter";
import { parseXmlDocument } from "./xmlParser";

describe("formatOoxmlCounter", () => {
  test("preserves the established counter representations", () => {
    expect(formatOoxmlCounter(7, "decimal")).toBe("7");
    expect(formatOoxmlCounter(7, "decimalZero")).toBe("07");
    expect(formatOoxmlCounter(7, "decimalZero3")).toBe("007");
    expect(formatOoxmlCounter(7, "decimalZero4")).toBe("0007");
    expect(formatOoxmlCounter(7, "decimalZero5")).toBe("00007");
    expect(formatOoxmlCounter(14, "upperRoman")).toBe("XIV");
    expect(formatOoxmlCounter(14, "lowerRoman")).toBe("xiv");
    expect(formatOoxmlCounter(27, "upperLetter")).toBe("AA");
    expect(formatOoxmlCounter(27, "lowerLetter")).toBe("aa");
    expect(formatOoxmlCounter(21, "ordinal")).toBe("21st");
    expect(formatOoxmlCounter(7, "bullet")).toBe("•");
    expect(formatOoxmlCounter(7, "none")).toBe("");
    expect(formatOoxmlCounter(7, "decimalEnclosedParen")).toBe("(7)");
    expect(formatOoxmlCounter(7, "numberInDash")).toBe("-7-");
    expect(formatOoxmlCounter(7, "japaneseCounting")).toBe("7");
  });

  test("formats Arabic alphabetic counters with Word's trailing non-joiner", () => {
    expect(formatOoxmlCounter(1, "arabicAlpha")).toBe("أ\u200c");
    expect(formatOoxmlCounter(2, "arabicAlpha")).toBe("ب\u200c");
    expect(formatOoxmlCounter(3, "arabicAlpha")).toBe("ت\u200c");
    expect(formatOoxmlCounter(28, "arabicAlpha")).toBe("ي\u200c");
    expect(formatOoxmlCounter(29, "arabicAlpha")).toBe("أأ\u200c");
    expect(formatOoxmlCounter(56, "arabicAlpha")).toBe("يي\u200c");
  });

  test("formats Arabic Abjad counters with Word's leading non-joiner", () => {
    expect(formatOoxmlCounter(1, "arabicAbjad")).toBe("\u200cأ");
    expect(formatOoxmlCounter(2, "arabicAbjad")).toBe("\u200cب");
    expect(formatOoxmlCounter(3, "arabicAbjad")).toBe("\u200cج");
    expect(formatOoxmlCounter(28, "arabicAbjad")).toBe("\u200cغ");
    expect(formatOoxmlCounter(29, "arabicAbjad")).toBe("\u200cأأ");
    expect(formatOoxmlCounter(56, "arabicAbjad")).toBe("\u200cغغ");
  });

  test("returns a string for every modeled OOXML number format", () => {
    for (const format of NUMBER_FORMAT_VALUES) {
      expect(typeof formatOoxmlCounter(7, format)).toBe("string");
    }
  });

  test("rejects non-finite internal counters", () => {
    expect(formatOoxmlCounter(Number.NaN, "decimal")).toBe("");
    expect(formatOoxmlCounter(Number.POSITIVE_INFINITY, "arabicAlpha")).toBe("");
  });

  test("bounds repeated-letter output for hostile external counter starts", () => {
    expect(formatOoxmlCounter(Number.MAX_SAFE_INTEGER, "lowerLetter")).toBe(
      String(Number.MAX_SAFE_INTEGER),
    );
    expect(formatOoxmlCounter(Number.MAX_SAFE_INTEGER, "arabicAlpha")).toBe(
      String(Number.MAX_SAFE_INTEGER),
    );
  });

  test("does not emit an isolated direction-control character for zero", () => {
    expect(formatOoxmlCounter(0, "arabicAlpha")).toBe("");
    expect(formatOoxmlCounter(0, "arabicAbjad")).toBe("");
  });
});

test("block parsing renders Arabic numbering formats through the shared formatter", () => {
  const numbering = parseNumbering(`
    <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:abstractNum w:abstractNumId="1">
        <w:lvl w:ilvl="0">
          <w:start w:val="28"/><w:numFmt w:val="arabicAlpha"/><w:lvlText w:val="%1."/>
        </w:lvl>
      </w:abstractNum>
      <w:abstractNum w:abstractNumId="2">
        <w:lvl w:ilvl="0">
          <w:start w:val="28"/><w:numFmt w:val="arabicAbjad"/><w:lvlText w:val="%1."/>
        </w:lvl>
      </w:abstractNum>
      <w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>
      <w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num>
    </w:numbering>
  `);
  const body = parseXmlDocument(`
    <w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:p><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:p>
      <w:p><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:p>
      <w:p><w:pPr><w:numPr><w:numId w:val="2"/></w:numPr></w:pPr></w:p>
      <w:p><w:pPr><w:numPr><w:numId w:val="2"/></w:numPr></w:pPr></w:p>
    </w:body>
  `);
  if (!body) {
    throw new Error("Failed to parse body XML");
  }

  const blocks = parseBlockContent(body, null, null, numbering, null, null);
  const markers = blocks.map((block) =>
    block.type === "paragraph" ? block.listRendering?.marker : undefined,
  );

  expect(markers).toEqual(["ي\u200c.", "أأ\u200c.", "\u200cغ.", "\u200cأأ."]);
});
