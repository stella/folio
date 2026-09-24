import { describe, expect, test } from "bun:test";

import type { ParagraphBlock } from "../types";
import { fixedCharWidth, withFakeTextMeasure } from "./__tests__/fakeTextMeasure";
import { measureParagraph } from "./measureParagraph";

const GLUE_CHARACTERS = ["\u00A0", "\u2007", "\u202F", "\u2060", "\uFEFF"];

/** Lay out text runs with 6 px glyphs and return each line's text. */
function lineTexts(texts: string[], width: number): string[] {
  let lines: string[] = [];
  withFakeTextMeasure(
    () => {
      const block: ParagraphBlock = {
        kind: "paragraph",
        id: "glue",
        runs: texts.map((text) => ({ kind: "text", text, fontFamily: "Arial", fontSize: 11 })),
      };
      lines = measureParagraph(block, width).lines.map((line) => {
        let text = "";
        for (let runIndex = line.fromRun; runIndex <= line.toRun; runIndex++) {
          const runText = texts[runIndex] ?? "";
          const start = runIndex === line.fromRun ? line.fromChar : 0;
          const end = runIndex === line.toRun ? line.toChar : runText.length;
          text += runText.slice(start, end);
        }
        return text;
      });
    },
    { charWidth: fixedCharWidth(6) },
  );
  return lines;
}

describe("measureParagraph — non-breaking glue across runs", () => {
  test.each(GLUE_CHARACTERS.map((glue) => [glue.codePointAt(0)?.toString(16), glue]))(
    "keeps U+%s glued to the next run's word",
    (_label, glue) => {
      const expected = ["aaaa ", `bbbb${glue}cc`];
      expect(lineTexts([`aaaa bbbb${glue}cc`], 60)).toEqual(expected);
      expect(lineTexts([`aaaa bbbb${glue}`, "cc"], 60)).toEqual(expected);
      expect(lineTexts(["aaaa bbbb", `${glue}cc`], 60)).toEqual(expected);
      expect(lineTexts(["aaaa bbbb", glue, "cc"], 60)).toEqual(expected);
    },
  );

  test.each(GLUE_CHARACTERS.map((glue) => [glue.codePointAt(0)?.toString(16), glue]))(
    "does not break an ideograph from U+%s in the next run",
    (_label, glue) => {
      const joined = lineTexts([`漢字漢字${glue}漢字漢字漢字`], 24);
      expect(joined).toEqual(["漢字漢", `字${glue}漢字`, "漢字漢字"]);
      expect(lineTexts(["漢字漢字", `${glue}漢字漢字漢字`], 24)).toEqual(joined);
      expect(lineTexts(["漢字漢字", glue, "漢字漢字漢字"], 24)).toEqual(joined);
    },
  );

  test("still breaks after a space that precedes a no-break space run", () => {
    expect(lineTexts(["aaaa bbbb ", "\u00A0cc"], 60)).toEqual(["aaaa bbbb ", "\u00A0cc"]);
  });
});
