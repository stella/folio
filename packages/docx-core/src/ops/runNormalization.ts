import type { ParagraphContent, Run } from "../model/document";
import { asParagraphContent, childNodes, rebuildNode } from "./leaves";
import { mergeRunContent, runsMergeable } from "./runMerge";

/**
 * Return the canonical segmentation of a paragraph's inline tree.
 *
 * Adjacent compatible runs are one run. Every other inline record is a
 * boundary, including range markers, fields, bookmarks and revision wrappers;
 * containers are normalised recursively without allowing a merge across their
 * edges.
 */
export const normalizeRunSegmentation = (
  content: readonly ParagraphContent[],
): ParagraphContent[] => {
  const normalized: ParagraphContent[] = [];

  for (const source of content) {
    const children = childNodes(source);
    const node =
      children === undefined || source.type === "run"
        ? source
        : rebuildNode(source, normalizeRunSegmentation(asParagraphContent(children)));
    if (node.type !== "run") {
      normalized.push(node as ParagraphContent);
      continue;
    }

    const previous = normalized.at(-1);
    if (previous?.type === "run" && runsMergeable(previous, node)) {
      normalized[normalized.length - 1] = {
        ...previous,
        content: mergeRunContent(previous.content, node.content),
      } satisfies Run;
      continue;
    }
    normalized.push(node);
  }

  return normalized;
};
