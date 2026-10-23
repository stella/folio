import { expect, test } from "bun:test";

import {
  fixedCharWidth,
  withFakeTextMeasure,
} from "../../layout-engine/measure/__tests__/fakeTextMeasure";
import { hashParagraphBlock } from "../../layout-engine/measure/cache";
import { measureParagraph } from "../../layout-engine/measure/measureParagraph";
import { toProseDoc } from "../../prosemirror/conversion/toProseDoc";
import type { Document } from "../../types/document";
import { toFlowBlocks } from "./toFlowBlocks";

test("document-default East Asian face reaches punctuation layout without mutating the source", () => {
  withFakeTextMeasure(
    () => {
      const keys = new Set<string>();
      for (const eastAsia of [undefined, "Arial", "MS Mincho"]) {
        const styles = {
          docDefaults: {
            rPr: { fontFamily: eastAsia === undefined ? {} : { eastAsia } },
          },
          styles: [],
        };
        const document = {
          package: {
            styles,
            document: {
              content: [
                {
                  type: "paragraph",
                  formatting: { alignment: "justify" },
                  content: [
                    {
                      type: "run",
                      formatting: { language: { eastAsia: "ja-JP" } },
                      content: [{ type: "text", text: "天地。" }],
                    },
                  ],
                },
              ],
            },
          },
        } satisfies Document;
        const original = JSON.stringify(document);
        const block = toFlowBlocks(toProseDoc(document, { styles }), {
          styles,
        }).at(0);
        if (block?.kind !== "paragraph") throw new Error("Expected paragraph");
        expect(block.attrs?.defaultEastAsiaFontFamily).toBe(eastAsia);
        expect(measureParagraph(block, 20).lines).toHaveLength(
          eastAsia === "MS Mincho" ? 1 : 2,
        );
        expect(JSON.stringify(document)).toBe(original);
        keys.add(hashParagraphBlock(block));
      }
      expect(keys.size).toBe(3);
    },
    { charWidth: fixedCharWidth(10) },
  );
});
