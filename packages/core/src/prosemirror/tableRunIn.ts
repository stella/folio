/**
 * A paragraph whose mark goes, right before a table.
 *
 * With its mark gone a paragraph runs on into what follows it. When that is a
 * table, its words run on into the table's first cell: they lead that cell's
 * first paragraph, which is the paragraph left, with its identity and
 * properties. What the removed paragraph's style lent its runs is re-read in
 * that paragraph's style, as in any join.
 *
 * A paragraph with nothing left to carry is simply removed by the resolvers;
 * one that ends a section keeps its place, as does one before a table whose
 * rows or cells are themselves pending.
 */

import { Fragment, type MarkType, type Node as PMNode } from "prosemirror-model";
import type { Transform } from "prosemirror-transform";

import {
  getProseParagraphPropertySourceToken,
  type ParagraphPropertySourceTransfer,
} from "../docx/paragraphPropertySource";
import { expectParagraphAttrs } from "./attrs";
import { rebaseParagraphRuns } from "./rebaseParagraphRunFormatting";
import { paragraphRunStyleContext, type RunStyleResolver } from "./runStyleFormatting";
import { holdsNoContent } from "./zeroWidthAnchors";

/** Whether a table's rows or cells are themselves pending insertion or deletion. */
export const tableHasPendingStructure = (table: PMNode): boolean => {
  let pending = false;
  table.descendants((child) => {
    pending ||=
      child.attrs["trIns"] != null ||
      child.attrs["trDel"] != null ||
      child.attrs["cellMarker"] != null;
    return !pending && !child.isTextblock;
  });
  return pending;
};

/** The first paragraph of a table's first cell, descending into nested tables. */
const runInTarget = (doc: PMNode, tablePos: number): { pos: number; node: PMNode } | null => {
  let node = doc.nodeAt(tablePos);
  let pos = tablePos;
  const table = node;
  const firstRow = table?.firstChild;
  if (!table || table.type.spec["tableRole"] !== "table" || !firstRow) return null;
  // A table whose rows or cells are themselves pending may not keep the cell
  // the words would run into: the paragraph keeps its own place instead.
  if (tableHasPendingStructure(table)) return null;
  while (node && node.type.name !== "paragraph") {
    const child: PMNode | null = node.firstChild;
    if (!child || child.isInline) return null;
    pos += 1;
    node = child;
  }
  return node ? { pos, node } : null;
};

type RunParagraphIntoTableOptions = {
  tr: Transform;
  paragraphPos: number;
  styleResolver: RunStyleResolver | null;
  /** Content a resolution in progress removes: it does not count as words left. */
  removedMark?: MarkType | undefined;
};

/**
 * Run the paragraph at `paragraphPos` on into the table right after it.
 * Returns where the paragraph that took its words now starts and the
 * paragraph-property source it hands over, or null when it did not run in and
 * the caller keeps its own handling.
 */
export const runParagraphIntoTable = ({
  tr,
  paragraphPos,
  styleResolver,
  removedMark,
}: RunParagraphIntoTableOptions): {
  position: number;
  transfer: ParagraphPropertySourceTransfer;
} | null => {
  const paragraph = tr.doc.nodeAt(paragraphPos);
  if (paragraph?.type.name !== "paragraph") return null;
  if (expectParagraphAttrs(paragraph)._sectionProperties !== undefined) return null;
  const left = removedMark
    ? paragraph.copy(
        Fragment.fromArray(paragraph.children.filter((child) => !removedMark.isInSet(child.marks))),
      )
    : paragraph;
  if (holdsNoContent(left)) return null;
  const tablePos = paragraphPos + paragraph.nodeSize;
  const target = runInTarget(tr.doc, tablePos);
  if (!target) return null;

  tr.insert(target.pos + 1, paragraph.content);
  if (styleResolver && paragraph.content.size > 0) {
    rebaseParagraphRuns({
      previousContext: paragraphRunStyleContext(paragraph, styleResolver),
      paragraphPosition: target.pos,
      range: { from: 0, to: paragraph.content.size },
      styleResolver,
      tr,
    });
  }
  tr.delete(paragraphPos, paragraphPos + paragraph.nodeSize);
  const displaced = getProseParagraphPropertySourceToken(paragraph);
  const selected = getProseParagraphPropertySourceToken(target.node);
  return {
    position: target.pos - paragraph.nodeSize,
    transfer: {
      displacedToken: typeof displaced === "string" ? displaced : null,
      selectedToken: typeof selected === "string" ? selected : null,
    },
  };
};

/** Whether resolving every change in `mode` removes this paragraph's mark. */
const markGoes = (paragraph: PMNode, mode: "accept" | "reject"): boolean => {
  const mark = expectParagraphAttrs(paragraph).pPrMark;
  if (!mark) return false;
  const added = mark.kind === "ins" || mark.kind === "moveTo";
  return added !== (mode === "accept");
};

/**
 * Before every change in a story is resolved: run each paragraph whose mark
 * goes and that sits right before a table on into that table. Returns each
 * paragraph that took words (its position after the step count given) and
 * the paragraph-property sources handed over.
 */
export const runParagraphsIntoTables = (
  tr: Transform,
  mode: "accept" | "reject",
  styleResolver: RunStyleResolver | null,
): {
  targets: { position: number; step: number }[];
  transfers: ParagraphPropertySourceTransfer[];
} => {
  const removedMark = tr.doc.type.schema.marks[mode === "accept" ? "deletion" : "insertion"];
  const tried = new Set<number>();
  const targets: { position: number; step: number }[] = [];
  const transfers: ParagraphPropertySourceTransfer[] = [];
  for (;;) {
    let candidate: number | null = null;
    tr.doc.descendants((node, pos, parent, index) => {
      if (node.type.name !== "paragraph") return !node.isTextblock;
      const next = parent?.maybeChild(index + 1);
      if (next?.type.spec["tableRole"] === "table" && markGoes(node, mode) && !tried.has(pos)) {
        candidate = pos;
      }
      return false;
    });
    if (candidate === null) return { targets, transfers };
    tried.add(candidate);
    const ranIn = runParagraphIntoTable({
      tr,
      paragraphPos: candidate,
      styleResolver,
      removedMark,
    });
    if (ranIn !== null) {
      targets.push({ position: ranIn.position, step: tr.steps.length });
      transfers.push(ranIn.transfer);
      tried.clear();
    }
  }
};
