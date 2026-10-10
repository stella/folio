import { panic } from "better-result";
import { Fragment } from "prosemirror-model";

import {
  createFolioAIEditSnapshotWithStyleResolver,
  sourceDocumentOf,
  styleResolverOf,
} from "../ai-edits/snapshot";
import type { FolioAIBlock, FolioAIEditSnapshot } from "../ai-edits/types";
import { alignFolioBlocks } from "../version-comparison";

type CreateInsertedResourceProjectionOptions = {
  base: FolioAIEditSnapshot;
  revised: FolioAIEditSnapshot;
};

/** Select resource nodes with their identities from the full revised snapshot. */
export const createInsertedResourceProjection = ({
  base,
  revised,
}: CreateInsertedResourceProjectionOptions) => {
  const document = sourceDocumentOf(revised);
  const insertions = alignFolioBlocks(base.blocks, revised.blocks).flatMap((event) => {
    if (event.type === "pair" || event.type === "baseOnly") return [];
    const anchor =
      revised.anchors[event.block.id] ?? panic("An inserted block lost its source anchor");
    const node =
      document.nodeAt(anchor.from) ?? panic("An inserted block lost its source paragraph");
    return [{ revisedBlockId: event.block.id, node }];
  });
  const snapshot = createFolioAIEditSnapshotWithStyleResolver(
    document.copy(Fragment.from(insertions.map(({ node }) => node))),
    styleResolverOf(revised),
  );
  if (snapshot.blocks.length !== insertions.length) {
    return panic("A redline resource projection must preserve one block per carried insertion");
  }
  return { snapshot, revisedBlockIds: insertions.map(({ revisedBlockId }) => revisedBlockId) };
};

/** Subset projection IDs never participate in rebinding the full revised snapshot. */
export const rebindInsertedResourceBlocks = (
  projection: ReturnType<typeof createInsertedResourceProjection>,
  rebound: FolioAIEditSnapshot,
) => {
  if (rebound.blocks.length !== projection.revisedBlockIds.length) {
    return panic("A redline resource import changed the carried insertion count");
  }
  const blocks = new Map<FolioAIBlock["id"], FolioAIBlock>();
  for (const [index, id] of projection.revisedBlockIds.entries()) {
    const block = rebound.blocks.at(index) ?? panic("A redline import lost a carried insertion");
    blocks.set(id, { ...block, id });
  }
  return blocks;
};
