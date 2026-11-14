/**
 * The run normal form of a paragraph: consecutive runs that can merge are
 * one run, at the paragraph's top level and inside hyperlinks, as the parser
 * consolidates them. A run holding nothing and every record other than a run
 * keep the runs beside them apart, and so does anything that makes two runs
 * differ (formatting, ids, preserved markup: `runsMergeable`).
 *
 * Every operation leaves the paragraphs it touched in this form, so the same
 * text, formatting and ids have one run structure whichever order of edits
 * reached them. Runs inside tracked changes, content controls and other
 * wrappers are left as they are.
 */
import type { Paragraph, Run } from "../model/document";
import { alikeDepth, type Gap, type InlineNode, runGaps } from "./leaves";
import { runsMergeable } from "./runMerge";

/** Where two consecutive runs that can merge meet, and how deep they merge. */
export type RunSeam = { at: Gap; depth: number };

const holdsPayload = (node: InlineNode): node is Run =>
  node.type === "run" && node.content.length > 0;

/** The seams a paragraph's normal form joins, in document order. */
export const runSeams = (paragraph: Paragraph): RunSeam[] => {
  const gaps = new Map(
    runGaps(paragraph.content).map(({ node, before }) => [node, before]),
  );
  const seams: RunSeam[] = [];
  const visit = (list: readonly InlineNode[]): void => {
    for (const [index, node] of list.entries()) {
      if (node.type === "hyperlink") visit(node.children);
      const left = list[index - 1];
      if (
        left === undefined ||
        !holdsPayload(left) ||
        !holdsPayload(node) ||
        !runsMergeable(left, node)
      )
        continue;
      const at = gaps.get(node);
      const depth = alikeDepth(left, node);
      if (at !== undefined && depth > 0) seams.push({ at, depth });
    }
  };
  visit(paragraph.content);
  return seams;
};
