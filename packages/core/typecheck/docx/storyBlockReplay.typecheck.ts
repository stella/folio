/** Replay consumes captured model records; unknown structural payloads stop at the boundary. */
import { buildStoryBlockReplay } from "../../src/docx/storyBlockReplay";
import type { BlockContent } from "../../src/types/document";

export const proveReplayBaselineContract = (
  blocks: readonly BlockContent[],
  untrusted: unknown,
) => {
  const source = { sourceXml: "", serializedXml: "", currentContent: blocks };
  buildStoryBlockReplay({ ...source, baselineContent: blocks });
  buildStoryBlockReplay({
    ...source,
    // @ts-expect-error An unknown JSON value is not a captured model block.
    baselineContent: [untrusted],
  });
  buildStoryBlockReplay({
    ...source,
    // @ts-expect-error Table rows must carry the source model's row and cell discriminators.
    baselineContent: [{ type: "table", rows: [{ cells: [{ content: [] }] }] }],
  });
};
