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

/**
 * Draws for properties whose laws start from a staged operation: the text
 * is never empty, so every insertion names an operation, and replacements
 * inside a block, which most documents can take, are drawn most often. Use
 * with `{ replaceIn: "blocksWithWords" }` so a replacement always finds a
 * word.
 */
export const stagingDrawArbitrary: fc.Arbitrary<OperationDraw> = fc.record({
  kind: fc.oneof(
    { weight: 3, arbitrary: fc.constant("replaceInBlock" as const) },
    { weight: 2, arbitrary: fc.constant("insertAfterBlock" as const) },
    { weight: 2, arbitrary: fc.constant("insertBeforeBlock" as const) },
    { weight: 1, arbitrary: fc.constant("deleteBlock" as const) },
  ),
  block: fc.nat(),
  word: fc.nat(),
  text: fc.constantFrom("revised", "the Customer", "Payment is due."),
});

type OperationForOptions = {
  /**
   * Which blocks a replacement draws from: every block a reader sees (a
   * block without words then names no operation), or only those with words.
   */
  replaceIn?: "everyBlock" | "blocksWithWords";
};

const wordsOf = (text: string): string[] =>
  text.split(/\s+/u).filter((word) => word !== "");

/** The operation a draw names over the blocks a reader of `reviewer` sees. */
export const operationFor = (
  reviewer: { getContent: () => FolioAIBlock[] },
  draw: OperationDraw,
  { replaceIn = "everyBlock" }: OperationForOptions = {},
): FolioDocumentOperation | null => {
  const read = reviewer
    .getContent()
    .filter((block) => block.kind !== "diagnostic");
  const blocks =
    draw.kind === "replaceInBlock" && replaceIn === "blocksWithWords"
      ? read.filter((block) => wordsOf(block.text).length > 0)
      : read;
  const block = blocks.at(draw.block % Math.max(blocks.length, 1));
  if (!block) return null;
  switch (draw.kind) {
    case "replaceInBlock": {
      const words = wordsOf(block.text);
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
