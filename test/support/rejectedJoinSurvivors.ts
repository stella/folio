import { isDeepStrictEqual } from "node:util";

import {
  documentStories,
  findStoryBody,
  paragraphIdsIn,
  storyBody,
  storyParagraphs,
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

/** A paragraph's properties: everything but its identity, content and mark. */
const propertiesOf = ({
  paraId: _paraId,
  textId: _textId,
  content: _content,
  pPrMark: _pPrMark,
  ...properties
}: Paragraph) => properties;

/**
 * `source` with the earlier paragraph of each survivor pair given the
 * properties of the later one, whose mark the join keeps with them; the
 * legacy reviewer, which keeps the earlier paragraph, then reads the joined
 * paragraph with the properties it has. `undefined` where every pair already
 * has the same properties.
 */
export const withSurvivorProperties = (
  source: Document,
  pairs: readonly { expectedId: string; survivorId: string }[],
): Document | undefined => {
  const paragraphs = new Map<string, Paragraph>();
  for (const story of documentStories(source))
    for (const { paragraph } of storyParagraphs(storyBody(source, story)))
      if (paragraph.paraId !== undefined)
        paragraphs.set(paragraph.paraId, paragraph);
  const replaced = new Map<string, Paragraph>();
  for (const { expectedId, survivorId } of pairs) {
    const earlier = paragraphs.get(expectedId);
    const later = paragraphs.get(survivorId);
    if (
      earlier === undefined ||
      later === undefined ||
      isDeepStrictEqual(propertiesOf(earlier), propertiesOf(later))
    )
      continue;
    replaced.set(expectedId, {
      ...propertiesOf(later),
      type: "paragraph",
      paraId: expectedId,
      ...(earlier.textId === undefined ? {} : { textId: earlier.textId }),
      content: earlier.content,
      ...(earlier.pPrMark === undefined ? {} : { pPrMark: earlier.pPrMark }),
    });
  }
  if (replaced.size === 0) return undefined;
  const visit = (value: unknown): unknown => {
    if (Array.isArray(value)) return value.map(visit);
    if (
      typeof value !== "object" ||
      value === null ||
      Object.getPrototypeOf(value) !== Object.prototype
    )
      return value;
    if (
      "type" in value &&
      value.type === "paragraph" &&
      "paraId" in value &&
      typeof value.paraId === "string"
    ) {
      const replacement = replaced.get(value.paraId);
      if (replacement !== undefined) return replacement;
    }
    return Object.fromEntries(
      Object.entries(value).map(([key, child]) => [key, visit(child)]),
    );
  };
  // SAFETY: only paragraphs are replaced, each by a paragraph.
  return visit(source) as Document;
};

/**
 * Rejecting the added mark of a paragraph that ends a local container and
 * holds nothing but added content has no paragraph to join it into there:
 * the paragraph stays, emptied and without its mark, as the container's
 * last. Where the block before it keeps its place, the legacy reviewer drops
 * it instead. Each id names such a paragraph,
 * proven on both rejected documents: the canonical one keeps it last in that
 * container, the legacy one has no such paragraph anywhere; every other
 * block stays exact.
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
        last.content.every(
          ({ type }) => type === "insertion" || type === "moveTo",
        ) &&
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

/** The paragraph a block content control's content ends with, nested ones included. */
const controlLastParagraph = (block: BlockContent): Paragraph | undefined => {
  if (block.type !== "blockSdt") return undefined;
  const last = block.content.at(-1);
  if (last === undefined) return undefined;
  return last.type === "paragraph" ? last : controlLastParagraph(last);
};

const deletedMark = (block: BlockContent | undefined): block is Paragraph =>
  block?.type === "paragraph" &&
  (block.pPrMark?.kind === "del" || block.pPrMark?.kind === "moveFrom");

type AbsorbedFollowersOptions = {
  source: Document;
  /** The canonical reviewer's resolved document. */
  actual: Document;
  /** The legacy reviewer's resolved document. */
  expected: Document;
  /** Rejecting resolves away added marks; accepting, deleted ones. */
  decision: "reject" | "accept";
};

/**
 * Resolving away the mark of the paragraph a block content control ends
 * with (rejecting an added one, accepting a deleted one) joins it with the
 * paragraph after the control: an empty one, which has nothing to bring into
 * the control, goes; the legacy reviewer keeps it. Each id names such a
 * paragraph, proven on both resolved documents: the canonical one has it no
 * more, the legacy one keeps it.
 */
export const resolvedAbsorbedFollowerIdentities = ({
  source,
  actual,
  expected,
  decision,
}: AbsorbedFollowersOptions): string[] => {
  const markGoes = decision === "reject" ? addedMark : deletedMark;
  const absorbed: string[] = [];
  for (const story of documentStories(source)) {
    const legacyIds = new Set(
      paragraphIdsIn(findStoryBody(expected, story)?.content ?? []),
    );
    const canonicalIds = new Set(
      paragraphIdsIn(findStoryBody(actual, story)?.content ?? []),
    );
    const walk = (blocks: readonly BlockContent[]): void => {
      for (const [index, block] of blocks.entries()) {
        for (const list of nestedLists(block)) walk(list);
        const follower = blocks[index + 1];
        if (
          markGoes(controlLastParagraph(block)) &&
          follower?.type === "paragraph" &&
          follower.content.length === 0 &&
          follower.pPrMark === undefined &&
          follower.sectionProperties === undefined &&
          typeof follower.paraId === "string" &&
          !canonicalIds.has(follower.paraId) &&
          legacyIds.has(follower.paraId)
        )
          absorbed.push(follower.paraId);
      }
    };
    walk(storyBody(source, story).content);
  }
  return absorbed;
};
