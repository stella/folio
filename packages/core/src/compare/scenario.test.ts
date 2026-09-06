import { describe, expect, test } from "bun:test";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { buildBodySequenceDocx } from "./__fixtures__/body-sequence";
import { buildNumberedListDocx } from "./__fixtures__/numbered-list";
import { applyEditScript } from "./scenario";

const blocksOf = async (buffer: ArrayBuffer) =>
  (await FolioDocxReviewer.fromBuffer(buffer)).getContent();

type MovedBlocksOptions = {
  base: ArrayBuffer;
  blockIndex: number;
  beforeBlockIndex: number;
};

const movedBlocks = async ({ base, blockIndex, beforeBlockIndex }: MovedBlocksOptions) => {
  const result = await applyEditScript(base, [
    { type: "moveParagraph", blockIndex, beforeBlockIndex },
  ]);
  if (result.isErr()) {
    throw result.error;
  }
  expect(result.value.unresolved).toEqual([]);
  return await blocksOf(result.value.buffer);
};

describe("move paragraph scenarios", () => {
  test.each([
    {
      source: "plain",
      sourceStyleId: undefined,
      destination: "plain",
      destinationStyleId: undefined,
    },
    {
      source: "plain",
      sourceStyleId: undefined,
      destination: "Heading1",
      destinationStyleId: "Heading1",
    },
    {
      source: "Heading1",
      sourceStyleId: "Heading1",
      destination: "plain",
      destinationStyleId: undefined,
    },
    {
      source: "Heading1",
      sourceStyleId: "Heading1",
      destination: "Heading1",
      destinationStyleId: "Heading1",
    },
  ] as const)(
    "preserve source style $source when moving before destination style $destination",
    async ({ sourceStyleId, destinationStyleId }) => {
      const base = await buildBodySequenceDocx([
        { kind: "paragraph", text: "Source", styleId: sourceStyleId },
        { kind: "paragraph", text: "Middle" },
        { kind: "paragraph", text: "Destination", styleId: destinationStyleId },
      ]);

      const blocks = await movedBlocks({ base, blockIndex: 0, beforeBlockIndex: 2 });

      expect(blocks.map(({ text, styleId }) => ({ text, styleId: styleId ?? null }))).toEqual([
        { text: "Middle", styleId: null },
        { text: "Source", styleId: sourceStyleId ?? null },
        { text: "Destination", styleId: destinationStyleId ?? null },
      ]);
    },
  );

  test.each([
    { source: "level 1", sourceLevel: 1, destination: "level 0", destinationLevel: 0 },
    { source: "unnumbered", sourceLevel: null, destination: "level 1", destinationLevel: 1 },
  ] as const)(
    "preserve source list state $source when moving before destination $destination",
    async ({ sourceLevel, destinationLevel }) => {
      const items = [
        { level: sourceLevel, text: "Source" },
        { level: 0, text: "Middle" },
        { level: destinationLevel, text: "Destination" },
      ];
      const base = await buildNumberedListDocx(items);

      const blocks = await movedBlocks({ base, blockIndex: 0, beforeBlockIndex: 2 });

      expect(blocks.map(({ text, listLevel }) => ({ text, listLevel: listLevel ?? null }))).toEqual(
        [
          { text: "Middle", listLevel: 0 },
          { text: "Source", listLevel: sourceLevel },
          { text: "Destination", listLevel: destinationLevel },
        ],
      );
    },
  );
});
