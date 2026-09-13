import { expect, test } from "bun:test";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { createEmptyDocument } from "../utils/createDocument";
import { createDocx } from "./rezip";

test.each([" PAGE ", "  PAGE   \\* MERGEFORMAT  ", ' MERGEFIELD "A B" \\* MERGEFORMAT '])(
  "preserves authored field instruction spacing across review saves: %s",
  async (instruction) => {
    const document = createEmptyDocument();
    document.package.document.content = [
      {
        type: "paragraph",
        content: [
          {
            type: "complexField",
            instruction,
            fieldCode: [],
            fieldResult: [{ type: "run", content: [{ type: "text", text: "1" }] }],
          },
        ],
      },
    ];
    let buffer = await createDocx(document);
    for (let pass = 0; pass < 2; pass++) {
      const reviewer = await FolioDocxReviewer.fromBuffer(buffer);
      const block = reviewer.toDocument().package.document.content.at(0);
      if (block?.type !== "paragraph") throw new Error("Expected field paragraph");
      const field = block.content.at(0);
      if (field?.type !== "complexField") throw new Error("Expected complex field");
      expect(field.instruction).toBe(instruction);
      buffer = await reviewer.toBuffer();
    }
  },
);
