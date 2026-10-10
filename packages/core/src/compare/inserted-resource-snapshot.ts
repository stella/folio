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
  const insertions = alignFolioBlocks(base.blocks, revised.blocks)
    .flatMap((event) => {
      if (event.type === "pair" || event.type === "baseOnly") return [];
      const anchor =
        revised.anchors[event.block.id] ?? panic("An inserted block lost its source anchor");
      const node =
        document.nodeAt(anchor.from) ?? panic("An inserted block lost its source paragraph");
      return [{ from: anchor.from, to: anchor.to, node }];
    })
    .toSorted((left, right) => left.from - right.from);
  const carriers: ((typeof insertions)[number] & { projectionFrom: number })[] = [];
  let projectionFrom = 0;
  for (const insertion of insertions) {
    // Copy each source subtree once, even when several inserted blocks belong to it.
    if (insertion.from < (carriers.at(-1)?.to ?? -1)) continue;
    carriers.push({ ...insertion, projectionFrom });
    projectionFrom += insertion.node.nodeSize;
  }
  // Carry every projected descendant by its position inside its maximal source
  // carrier. Subset-generated IDs cannot identify duplicate or positional IDs.
  const revisedIdsByProjectionPosition = new Map<number, FolioAIBlock["id"]>();
  let carrierIndex = 0;
  for (const block of revised.blocks) {
    const anchor = revised.anchors[block.id];
    if (anchor === undefined) continue;
    while (anchor.from >= (carriers.at(carrierIndex)?.to ?? Number.POSITIVE_INFINITY)) {
      carrierIndex++;
    }
    const carrier = carriers.at(carrierIndex);
    if (carrier === undefined || anchor.from < carrier.from || anchor.to > carrier.to) continue;
    revisedIdsByProjectionPosition.set(
      carrier.projectionFrom + anchor.from - carrier.from,
      block.id,
    );
  }
  const snapshot = createFolioAIEditSnapshotWithStyleResolver(
    document.copy(Fragment.from(carriers.map(({ node }) => node))),
    styleResolverOf(revised),
  );
  if (snapshot.blocks.length !== revisedIdsByProjectionPosition.size) {
    return panic("A redline resource projection must preserve its maximal carriers' blocks");
  }
  const revisedBlockIds = snapshot.blocks.map((block) => {
    const anchor =
      snapshot.anchors[block.id] ?? panic("A projected resource block lost its anchor");
    return (
      revisedIdsByProjectionPosition.get(anchor.from) ??
      panic("A projected resource block lost its carried source identity")
    );
  });
  return { snapshot, revisedBlockIds };
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
