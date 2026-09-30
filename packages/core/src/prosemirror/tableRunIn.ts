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

import { Fragment, Slice, type MarkType, type Node as PMNode } from "prosemirror-model";
import { ReplaceAroundStep, StepMap, type Transform } from "prosemirror-transform";

import {
  getProseParagraphPropertySourceToken,
  type ParagraphPropertySourceTransfer,
  recreateProseNodeWithParagraphPropertySource as rebuild,
} from "../docx/paragraphPropertySource";
import { expectParagraphAttrs } from "./attrs";
import { rebaseParagraphRuns } from "./rebaseParagraphRunFormatting";
import { paragraphRunStyleContext, type RunStyleResolver } from "./runStyleFormatting";
import { holdsNoContent } from "./zeroWidthAnchors";

/**
 * Whether the cell a paragraph's words would run into is itself pending: its
 * row or the cell (through nested tables' first cells) is inserted, deleted
 * or merged, so resolving may not keep it. Pending rows or cells elsewhere in
 * the table leave that cell where it is.
 */
export const runInCellPending = (table: PMNode): boolean => {
  let node: PMNode | null = table;
  while (node?.type.spec["tableRole"] === "table") {
    const row: PMNode | null = node.firstChild;
    if (!row) return false;
    if (row.attrs["trIns"] != null || row.attrs["trDel"] != null) return true;
    const cell: PMNode | null = row.firstChild;
    if (!cell) return false;
    if (cell.attrs["cellMarker"] != null) return true;
    node = cell.firstChild;
  }
  return false;
};

/** The first paragraph of a table's first cell, descending into nested tables. */
const runInTarget = (doc: PMNode, tablePos: number): { pos: number; node: PMNode } | null => {
  let node = doc.nodeAt(tablePos);
  let pos = tablePos;
  const table = node;
  const firstRow = table?.firstChild;
  if (!table || table.type.spec["tableRole"] !== "table" || !firstRow) return null;
  // A pending row or cell may not keep the cell the words would run into:
  // the paragraph keeps its own place instead.
  if (runInCellPending(table)) return null;
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
  map: StepMap;
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

  // One step that keeps the words where they are and moves the table's
  // opening (table, row, cell, the cell's first paragraph) in front of them,
  // so every position in the words still maps to itself.
  const $target = tr.doc.resolve(target.pos + 1);
  const containerDepth = tr.doc.resolve(paragraphPos).depth;
  let opening = Fragment.empty;
  for (let depth = $target.depth; depth > containerDepth; depth--) {
    opening = Fragment.from(rebuild($target.node(depth), { content: opening }));
  }
  const levels = $target.depth - containerDepth;
  tr.step(
    new ReplaceAroundStep(
      paragraphPos,
      target.pos + 1,
      paragraphPos + 1,
      paragraphPos + paragraph.nodeSize - 1,
      new Slice(opening, 0, levels),
      levels,
      true,
    ),
  );
  const position = paragraphPos + levels - 1;
  if (styleResolver && paragraph.content.size > 0) {
    rebaseParagraphRuns({
      previousContext: paragraphRunStyleContext(paragraph, styleResolver),
      paragraphPosition: position,
      range: { from: 0, to: paragraph.content.size },
      styleResolver,
      tr,
    });
  }
  const displaced = getProseParagraphPropertySourceToken(paragraph);
  const selected = getProseParagraphPropertySourceToken(target.node);
  // How positions read across the move: the paragraph's opening is the cell
  // paragraph's, the table, row and cell openings are new in front of it, and
  // the paragraph's close and the table's openings after its words are gone.
  const closing = paragraphPos + paragraph.nodeSize - 1;
  const map = new StepMap([paragraphPos, 0, levels - 1, closing, target.pos + 1 - closing, 0]);
  return {
    position,
    map,
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
 * Whether resolving every change in `mode` runs `paragraph` on into `next`:
 * its mark goes and `next` is a table.
 */
export const runsIntoFollowingTable = (
  paragraph: PMNode,
  next: PMNode | null | undefined,
  mode: "accept" | "reject",
): boolean =>
  paragraph.type.name === "paragraph" &&
  next?.type.spec["tableRole"] === "table" &&
  markGoes(paragraph, mode);

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
  maps: StepMap[];
  transfers: ParagraphPropertySourceTransfer[];
} => {
  const removedMark = tr.doc.type.schema.marks[mode === "accept" ? "deletion" : "insertion"];
  const tried = new Set<number>();
  const targets: { position: number; step: number }[] = [];
  const transfers: ParagraphPropertySourceTransfer[] = [];
  const maps: StepMap[] = [];
  for (;;) {
    let candidate: number | null = null;
    tr.doc.descendants((node, pos, parent, index) => {
      if (node.type.name !== "paragraph") return !node.isTextblock;
      const next = parent?.maybeChild(index + 1);
      if (runsIntoFollowingTable(node, next, mode) && !tried.has(pos)) {
        candidate = pos;
      }
      return false;
    });
    if (candidate === null) return { targets, maps, transfers };
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
      maps.push(ranIn.map);
      tried.clear();
    }
  }
};
