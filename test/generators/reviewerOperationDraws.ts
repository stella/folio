/** One public operation drawn over the blocks a reviewer's reader sees. */
import fc from "fast-check";

import type { FolioAIBlock } from "../../packages/core/src/ai-edits/types";
import type { FolioDocumentOperation } from "../../packages/core/src/document-operations";

/** What the operation is built from: one draw per choice it makes. */
export type OperationDraw = {
  kind:
    "replaceInBlock" | "insertAfterBlock" | "insertBeforeBlock" | "deleteBlock";
  block: number;
  word: number;
  text: string;
};

export const operationDrawArbitrary: fc.Arbitrary<OperationDraw> = fc.record({
  kind: fc.constantFrom(
    "replaceInBlock",
    "insertAfterBlock",
    "insertBeforeBlock",
    "deleteBlock",
  ),
  block: fc.nat(),
  word: fc.nat(),
  text: fc.constantFrom("revised", "the Customer", "Payment is due.", ""),
});

/** The operation a draw names over the blocks a reader of `reviewer` sees. */
export const operationFor = (
  reviewer: { getContent: () => FolioAIBlock[] },
  draw: OperationDraw,
): FolioDocumentOperation | null => {
  const blocks = reviewer
    .getContent()
    .filter((block) => block.kind !== "diagnostic");
  const block = blocks.at(draw.block % Math.max(blocks.length, 1));
  if (!block) return null;
  switch (draw.kind) {
    case "replaceInBlock": {
      const words = block.text.split(/\s+/u).filter((word) => word !== "");
      const find = words.at(draw.word % Math.max(words.length, 1));
      return find === undefined
        ? null
        : {
            id: "op",
            type: "replaceInBlock",
            blockId: block.id,
            find,
            replace: draw.text,
          };
    }
    case "insertAfterBlock":
    case "insertBeforeBlock":
      return draw.text === ""
        ? null
        : { id: "op", type: draw.kind, blockId: block.id, text: draw.text };
    case "deleteBlock":
      return { id: "op", type: "deleteBlock", blockId: block.id };
    default: {
      const unreachable: never = draw.kind;
      throw new Error(`Unknown operation draw ${String(unreachable)}`);
    }
  }
};
