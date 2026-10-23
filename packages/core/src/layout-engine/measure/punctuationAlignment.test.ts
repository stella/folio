import { describe, expect, test } from "bun:test";

import type { ParagraphBlock } from "../types";
import {
  fixedCharWidth,
  withFakeTextMeasure,
} from "./__tests__/fakeTextMeasure";
import { measureParagraph } from "./measureParagraph";

const ALIGNMENTS = [undefined, "left", "center", "right", "justify"] as const;
const LINE_WIDTH = 20;

describe("punctuation overflow follows paragraph alignment", () => {
  test.each(ALIGNMENTS)("%s alignment", (alignment) => {
    withFakeTextMeasure(
      () => {
        for (const punctuation of ["。", "”", ",", "."]) {
          for (const prefix of ["天地", "AB"]) {
            for (const overflowPunctuation of [undefined, true, false]) {
              const block = {
                kind: "paragraph",
                id: "punctuation-alignment",
                runs: [
                  {
                    kind: "text",
                    text: `${prefix}${punctuation}`,
                    language: { val: "ja-JP", eastAsia: "ja-JP" },
                  },
                ],
                attrs: { alignment, overflowPunctuation },
              } as const satisfies ParagraphBlock;

              const measure = measureParagraph(block, LINE_WIDTH);
              if (alignment === "justify" && overflowPunctuation !== false) {
                expect(measure.lines).toHaveLength(1);
                expect(measure.lines.at(0)?.width).toBe(30);
                continue;
              }

              expect(measure.lines.length).toBeGreaterThan(1);
              for (const line of measure.lines) {
                expect(line.width).toBeLessThanOrEqual(LINE_WIDTH);
              }
            }
          }
        }
      },
      { charWidth: fixedCharWidth(10) },
    );
  });
});
