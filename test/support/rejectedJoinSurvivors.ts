import {
  documentStories,
  findStoryBody,
  paragraphIdsIn,
  storyBody,
} from "@stll/docx-core/ops";

import type {
  BlockContent,
  Document,
  Paragraph,
} from "../../packages/docx-core/src/model/document";

const addedMark = (block: BlockContent | undefined): block is Paragraph =>
  block?.type === "paragraph" &&
  (block.pPrMark?.kind === "ins" || block.pPrMark?.kind === "moveTo");

/** The block lists nested in a block: a content control's or cell's blocks. */
const nestedLists = (block: BlockContent): readonly BlockContent[][] => {
  switch (block.type) {
    case "blockSdt":
    case "blockCustomXml":
      return [block.content];
    case "table":
      return block.rows.flatMap((row) => row.cells.map((cell) => cell.content));
    default:
      return [];
  }
};

type RejectedJoinSurvivorsOptions = {
  source: Document;
  /** The canonical reviewer's rejected document. */
  actual: Document;
  /** The legacy reviewer's rejected document. */
  expected: Document;
};

/**
 * Removing a paragraph mark leaves the paragraph after it: rejecting an added
 * mark joins the paragraph into the next one, which keeps its identity. Where
 * that next paragraph's own added mark stays, because it ends a local
 * container or a block it cannot join follows it, the legacy reviewer keeps
 * the earlier paragraph's identity instead. Each pair names the id the legacy
 * reviewer reports and the id the join keeps, proven on both rejected
 * documents; every other id stays exact.
 */
export const rejectedJoinSurvivorIdentities = ({
  source,
  actual,
  expected,
}: RejectedJoinSurvivorsOptions) => {
  const pairs: { expectedId: string; survivorId: string }[] = [];
  for (const story of documentStories(source)) {
    const legacyIds = new Set(
      paragraphIdsIn(findStoryBody(expected, story)?.content ?? []),
    );
    const canonicalIds = new Set(
      paragraphIdsIn(findStoryBody(actual, story)?.content ?? []),
    );
    const walk = (blocks: readonly BlockContent[], local: boolean): void => {
      for (const [index, block] of blocks.entries()) {
        for (const list of nestedLists(block)) walk(list, true);
        const next = blocks.at(index + 1);
        const after = blocks.at(index + 2);
        const nextMarkStays =
          after === undefined
            ? local
            : after.type !== "paragraph" && after.type !== "table";
        if (
          !addedMark(block) ||
          !addedMark(next) ||
          !nextMarkStays ||
          typeof block.paraId !== "string" ||
          typeof next.paraId !== "string" ||
          !legacyIds.has(block.paraId) ||
          legacyIds.has(next.paraId) ||
          canonicalIds.has(block.paraId) ||
          !canonicalIds.has(next.paraId)
        )
          continue;
        pairs.push({ expectedId: block.paraId, survivorId: next.paraId });
      }
    };
    walk(storyBody(source, story).content, false);
  }
  return pairs;
};

/**
 * Rejecting the added mark of an empty paragraph that ends a local container
 * has no paragraph to join it into there: the paragraph stays, without its
 * mark, as the container's last. Where the block before it keeps its place,
 * the legacy reviewer drops it instead. Each
 * id names such a paragraph, proven on both rejected documents: the
 * canonical one keeps it last in that container, the legacy one has no such
 * paragraph anywhere; every other block stays exact.
 */
export const rejectedContainerEndIdentities = ({
  source,
  actual,
  expected,
}: RejectedJoinSurvivorsOptions): string[] => {
  const kept: string[] = [];
  for (const story of documentStories(source)) {
    const legacyIds = new Set(
      paragraphIdsIn(findStoryBody(expected, story)?.content ?? []),
    );
    const canonicalLastIds = new Set<string>();
    const collect = (blocks: readonly BlockContent[], local: boolean) => {
      const last = blocks.at(-1);
      if (
        local &&
        last?.type === "paragraph" &&
        last.pPrMark === undefined &&
        typeof last.paraId === "string"
      )
        canonicalLastIds.add(last.paraId);
      for (const block of blocks)
        for (const list of nestedLists(block)) collect(list, true);
    };
    collect(findStoryBody(actual, story)?.content ?? [], false);
    const walk = (blocks: readonly BlockContent[], local: boolean): void => {
      const last = blocks.at(-1);
      const previous = blocks.at(-2);
      if (
        local &&
        previous !== undefined &&
        !addedMark(previous) &&
        addedMark(last) &&
        last.content.length === 0 &&
        typeof last.paraId === "string" &&
        canonicalLastIds.has(last.paraId) &&
        !legacyIds.has(last.paraId)
      )
        kept.push(last.paraId);
      for (const block of blocks)
        for (const list of nestedLists(block)) walk(list, true);
    };
    walk(storyBody(source, story).content, false);
  }
  return kept;
};
