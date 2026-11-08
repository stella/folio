import { expect, test } from "bun:test";

import type { FlowBlock } from "../../layout-engine/types";
import { indexBlocksById } from "./blockIndex";

const paragraph = (id: string): FlowBlock => ({
  kind: "paragraph",
  id,
  runs: [],
  attrs: {},
});

test("each block id finds the index a search from the start finds", () => {
  const blocks = ["a", "b", "a", "c", "b"].map(paragraph);
  const indices = indexBlocksById(blocks);
  for (const block of blocks)
    expect(indices.get(block.id)).toBe(
      blocks.findIndex((candidate) => candidate.id === block.id),
    );
  expect(indices.get("missing")).toBeUndefined();
});
