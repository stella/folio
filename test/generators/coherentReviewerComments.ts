/** Coherent comments for the raw review generator's fixed blank-root template. */
import { panic } from "better-result";

import type {
  BlockContent,
  Document,
} from "../../packages/docx-core/src/model/document";
import { withBodyContent } from "../../packages/docx-core/src/ops/blocks";
import {
  commentAnchorsIn,
  commentDocumentIssue,
  freshCommentId,
} from "../../packages/docx-core/src/ops/comments";
import {
  formatParaId,
  freshLeasedIds,
  ID_LEASE_KINDS,
  leasedIdCensus,
} from "../../packages/docx-core/src/ops/idLease";
import {
  asParagraphContent,
  childNodes,
  type CommentAnchorCarrier,
  type InlineNode,
  isCommentAnchor,
  rebuildNode,
} from "../../packages/docx-core/src/ops/leaves";
import { OP_STORIES } from "../../packages/docx-core/src/ops/types";

/**
 * The operation generator draws independent anchor atoms, not comment graphs.
 * Coherent inputs stay unchanged. Repair supports only the raw generator's
 * blank root; package/reply fixtures use their own coherent construction.
 */
export const withCoherentSyntheticReviewComments = (
  document: Document,
): Document => {
  if (commentDocumentIssue(document, "all") === undefined) return document;
  const comments = document.package.document.comments;
  const root = comments?.at(0);
  const body = root?.content.at(0);
  if (
    comments?.length !== 1 ||
    root === undefined ||
    typeof root.author !== "string" ||
    root.parentId !== undefined ||
    root.content.length !== 1 ||
    body?.paraId === undefined ||
    body.content.length !== 0
  )
    return panic(
      "Reviewer comment construction requires the raw blank-root fixture.",
    );

  const anchors = commentAnchorsIn(document);
  if (
    anchors.some(
      ({ story, betweenBlocks }) => story !== OP_STORIES.MAIN || betweenBlocks,
    )
  )
    return panic(
      "The raw review generator's comment atoms must be inline in the main story.",
    );

  let allocated = document;
  const fresh = () => {
    const id = freshCommentId(allocated).unwrap();
    const occupied = leasedIdCensus(allocated).paragraph;
    const paragraphId = freshLeasedIds({
      lease: undefined,
      kind: ID_LEASE_KINDS.PARAGRAPH,
      count: 1,
      taken: (candidate) => occupied.has(candidate),
    }).at(0);
    if (paragraphId === undefined)
      return panic(
        "The synthetic comment paragraph identity space is exhausted.",
      );
    allocated = {
      ...allocated,
      package: {
        ...allocated.package,
        document: {
          ...allocated.package.document,
          comments: [
            ...(allocated.package.document.comments ?? []),
            {
              id,
              author: root.author,
              content: [
                {
                  type: "paragraph",
                  paraId: formatParaId(paragraphId),
                  content: [],
                },
              ],
            },
          ],
        },
      },
    };
    return id;
  };

  const groups = new Map<
    number,
    { index: number; node: CommentAnchorCarrier }[]
  >();
  const planned = anchors.map(({ node }, index) => {
    const group = groups.get(node.id);
    const occurrence = { index, node };
    if (group === undefined) groups.set(node.id, [occurrence]);
    else group.push(occurrence);
    return [node];
  });
  for (const [originalId, group] of groups) {
    let existing = originalId === root.id ? originalId : undefined;
    const binding = () => {
      if (existing === undefined) return fresh();
      const id = existing;
      existing = undefined;
      return id;
    };
    const starts = group.filter(
      ({ node }) => node.type === "commentRangeStart",
    );
    const ends = group.filter(({ node }) => node.type === "commentRangeEnd");
    const references = group.filter(
      ({ node }) => node.type === "commentReference",
    );
    const start = starts.at(0),
      end = ends.at(0);
    if (
      starts.length === 1 &&
      ends.length === 1 &&
      start !== undefined &&
      end !== undefined &&
      start.index < end.index
    ) {
      const id = binding();
      planned[start.index] = [{ ...start.node, id }];
      planned[end.index] = [{ ...end.node, id }];
      const firstReference = references.at(0);
      if (firstReference !== undefined)
        planned[firstReference.index] = [{ ...firstReference.node, id }];
      for (const reference of references.slice(1))
        planned[reference.index] = [{ ...reference.node, id: fresh() }];
      continue;
    }
    for (const { index, node } of group) {
      const id = binding();
      switch (node.type) {
        case "commentReference":
          planned[index] = [{ ...node, id }];
          break;
        case "commentRangeStart":
          planned[index] = [
            { ...node, id },
            { type: "commentRangeEnd", id },
          ];
          break;
        case "commentRangeEnd":
          planned[index] = [
            { type: "commentRangeStart", id },
            { ...node, id },
          ];
          break;
        default: {
          const unreachable: never = node;
          return unreachable;
        }
      }
    }
  }

  // Both passes walk the original source, before any new mate is inserted.
  // Reused constant atoms are distinct occurrences even when they share an object.
  let occurrence = 0;
  const inline = (nodes: readonly InlineNode[]): InlineNode[] =>
    nodes.flatMap((node): InlineNode[] => {
      if (isCommentAnchor(node)) {
        const original = anchors.at(occurrence);
        const replacement = planned.at(occurrence++);
        if (original?.node !== node || replacement === undefined)
          return panic(
            "Reviewer comment traversal lost its source occurrence.",
          );
        return replacement;
      }
      const children = childNodes(node);
      return [
        children === undefined ? node : rebuildNode(node, inline(children)),
      ];
    });
  const blocks = (content: readonly BlockContent[]): BlockContent[] =>
    content.map((block): BlockContent => {
      switch (block.type) {
        case "paragraph":
          return {
            ...block,
            content: asParagraphContent(inline(block.content)),
          };
        case "table":
          return {
            ...block,
            rows: block.rows.map((row) => ({
              ...row,
              cells: row.cells.map((cell) => ({
                ...cell,
                content: blocks(cell.content),
              })),
            })),
          };
        case "blockSdt":
        case "blockCustomXml":
          return { ...block, content: blocks(block.content) };
        default:
          return block;
      }
    });
  const content = blocks(document.package.document.content);
  if (occurrence !== anchors.length)
    return panic(
      "Reviewer comment construction left source occurrences unvisited.",
    );
  return {
    ...allocated,
    package: {
      ...allocated.package,
      document: withBodyContent(allocated.package.document, content),
    },
  };
};
