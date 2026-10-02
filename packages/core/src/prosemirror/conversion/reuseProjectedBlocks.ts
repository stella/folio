import { panic } from "better-result";
import type { Node as PMNode } from "prosemirror-model";

import type { BlockContent, Document, Paragraph } from "../../types/document";
import {
  getDocumentParagraphPropertySourceContract,
  getParagraphPropertySourceToken,
  getProseParagraphPropertySourceToken,
  visitDocumentStoryParagraphs,
} from "../../docx/paragraphPropertySource";
import { visitParagraphRuns } from "../../docx/paragraphTraversal";
import { visitCommentMarkers } from "../../docx/commentAnchorIndex";
import { toProseDoc } from "./toProseDoc";

const projectedIndexes = new WeakMap<PMNode, Map<string, PMNode | null>>();

// ProseMirror nodes are immutable. Index each tree in one bottom-up pass,
// rather than rescanning every table/wrapper's paragraph descendants.
const uniqueProjectedBlocks = (document: PMNode): Map<string, PMNode | null> => {
  const cached = projectedIndexes.get(document);
  if (cached) return cached;
  const blocks = new Map<string, PMNode | null>();
  const collect = (node: PMNode): string | undefined => {
    const token =
      node.type.name === "paragraph" ? getProseParagraphPropertySourceToken(node) : undefined;
    let firstToken = typeof token === "string" ? token : undefined;
    // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
    node.forEach((child) => {
      const childToken = collect(child);
      firstToken ??= childToken;
    });
    const identityToken = node.type.name === "paragraph" ? token : firstToken;
    switch (node.type.name) {
      case "paragraph":
      case "table":
      case "blockSdt":
      case "blockCustomXml":
        if (typeof identityToken === "string") {
          const identity = `${node.type.name}:${identityToken}`;
          blocks.set(identity, blocks.has(identity) ? null : node);
        }
        break;
    }
    return firstToken;
  };
  collect(document);
  projectedIndexes.set(document, blocks);
  return blocks;
};

/** A paragraph split into sibling editor nodes needs a group comparison. */
const hasSiblingProjection = (paragraph: Paragraph): boolean => {
  let split = false;
  visitParagraphRuns(paragraph, (run) => {
    split ||= run.content.some(
      (content) =>
        (content.type === "break" && content.breakType === "page") ||
        (content.type === "shape" && content.shape.shapeType === "textBox"),
    );
  });
  return split;
};

type ModelBlockIndex = {
  sources: Map<string, BlockContent | null>;
  identities: WeakMap<BlockContent, string>;
};
const sourceIndexes = new WeakMap<Document, { projection: PMNode; index: ModelBlockIndex }>();

// Editor marks are paragraph-local; extraction repairs a surviving half of a
// cross-paragraph comment after deletion. Projection equality cannot see that
// boundary repair, so it must agree before an authored record can be retained.
const commentBoundaries = (block: BlockContent): string[] => {
  const boundaries: string[] = [];
  visitCommentMarkers([block], ({ item }) => {
    boundaries.push(`${item.type}:${item.id}`);
  });
  return boundaries;
};

const sameCommentBoundaries = (source: BlockContent, converted: BlockContent): boolean => {
  const expected = commentBoundaries(source);
  const actual = commentBoundaries(converted);
  return (
    expected.length === actual.length && expected.every((marker, index) => marker === actual[index])
  );
};

const collectModelBlocks = (content: BlockContent[]): ModelBlockIndex => {
  const sources = new Map<string, BlockContent | null>();
  const identities = new WeakMap<BlockContent, string>();
  const collect = (blocks: BlockContent[]): string | undefined => {
    let firstToken: string | undefined;
    for (const block of blocks) {
      let token: string | undefined;
      switch (block.type) {
        case "paragraph":
          token = getParagraphPropertySourceToken(block);
          // A text box may own the first token when its anchor is unbound.
          if (token === undefined) {
            visitDocumentStoryParagraphs([block], (paragraph) => {
              token ??= getParagraphPropertySourceToken(paragraph);
            });
          }
          break;
        case "table":
          for (const row of block.rows) {
            for (const cell of row.cells) {
              const childToken = collect(cell.content);
              token ??= childToken;
            }
          }
          break;
        case "blockSdt":
        case "blockCustomXml":
          token = collect(block.content);
          break;
        case "preservedBlock":
        case "bookmarkStart":
        case "bookmarkEnd":
          break;
        default: {
          const unsupported: never = block;
          panic(`Unsupported projection source block: ${JSON.stringify(unsupported)}`);
        }
      }
      firstToken ??= token;
      if (token === undefined) continue;
      const identity = `${block.type}:${token}`;
      identities.set(block, identity);
      sources.set(identity, sources.has(identity) ? null : block);
    }
    return firstToken;
  };
  collect(content);
  return { sources, identities };
};

type ReuseProjectedBlocksOptions = {
  blocks: BlockContent[];
  projected: PMNode;
  base: Document;
  sourceProjection: PMNode | undefined;
  stripSuggested: (projection: PMNode) => PMNode;
};

export const reuseProjectedBlocks = ({
  blocks,
  projected,
  base,
  sourceProjection: validatedProjection,
  stripSuggested,
}: ReuseProjectedBlocksOptions): BlockContent[] => {
  if (!getDocumentParagraphPropertySourceContract(base)) {
    return blocks;
  }
  const sourceProjection = validatedProjection ?? toProseDoc(base);
  const expected = uniqueProjectedBlocks(sourceProjection);
  const strippedProjection = stripSuggested(sourceProjection);
  const stripped =
    strippedProjection === sourceProjection ? expected : uniqueProjectedBlocks(strippedProjection);
  const actual = uniqueProjectedBlocks(projected);
  // A validated projection also versions its mutable source model and tokens.
  // Rebuild this index only when that projection changes.
  const cachedSources = sourceIndexes.get(base);
  const sourceIndex =
    cachedSources?.projection === sourceProjection
      ? cachedSources.index
      : collectModelBlocks(base.package.document.content);
  sourceIndexes.set(base, { projection: sourceProjection, index: sourceIndex });
  const { sources } = sourceIndex;
  const { identities } = collectModelBlocks(blocks);

  const reuseBlocks = (content: BlockContent[]): BlockContent[] =>
    content.map((block) => {
      const identity = identities.get(block);
      const source = identity === undefined ? undefined : sources.get(identity);
      const expectedBlock = identity === undefined ? undefined : expected.get(identity);
      const strippedBlock = identity === undefined ? undefined : stripped.get(identity);
      const actualBlock = identity === undefined ? undefined : actual.get(identity);
      if (
        source &&
        source.type === block.type &&
        expectedBlock &&
        actualBlock &&
        expectedBlock === strippedBlock &&
        (source.type !== "paragraph" ||
          (block.type === "paragraph" &&
            block.paraId === source.paraId &&
            !hasSiblingProjection(source))) &&
        expectedBlock.eq(actualBlock) &&
        sameCommentBoundaries(source, block)
      ) {
        return source;
      }
      switch (block.type) {
        case "paragraph":
          return block;
        case "table":
          return {
            ...block,
            rows: block.rows.map((row) => ({
              ...row,
              cells: row.cells.map((cell) => ({ ...cell, content: reuseBlocks(cell.content) })),
            })),
          };
        case "blockSdt":
        case "blockCustomXml":
          return { ...block, content: reuseBlocks(block.content) };
        case "preservedBlock":
        case "bookmarkStart":
        case "bookmarkEnd":
          return block;
        default: {
          const unsupported: never = block;
          panic(`Unsupported projection reuse block: ${JSON.stringify(unsupported)}`);
        }
      }
    });
  return reuseBlocks(blocks);
};
