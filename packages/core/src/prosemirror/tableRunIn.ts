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
 * first row or cell may not stay. Rejecting every change gives the document
 * before back, so a deleted row stays for it. A table that resolving every
 * change removes never stood in the way: the paragraph runs on into what follows
 * that table.
 */

import { Fragment, type Node as PMNode, Slice } from "prosemirror-model";
import {
  ReplaceAroundStep,
  ReplaceStep,
  StepMap,
  type Transform,
} from "prosemirror-transform";

import { paragraphMarkOutcome } from "@stll/docx-core/ops";

import {
  getProseParagraphPropertySourceToken,
  type ParagraphPropertySourceTransfer,
  recreateProseNodeWithParagraphPropertySource as rebuild,
} from "../docx/paragraphPropertySource";
import { expectParagraphAttrs } from "./attrs";
import { expectCommentReferenceAttrs } from "./commentReferenceAttrs";
import { COMMENT_REFERENCE_NODE_NAME } from "./extensions/nodes/CommentReferenceExtension";
import {
  paragraphRemainsOf,
  REVISION_MARK_NAMES,
  revisionLayerRemovesContent,
} from "./keptRangeBoundaries";
import { rebaseParagraphRuns } from "./rebaseParagraphRunFormatting";
import {
  paragraphRunStyleContext,
  type RunStyleResolver,
} from "./runStyleFormatting";
import { trackedRevisionPathOf } from "./trackedRevisionPath";
import { paragraphContentOf } from "./zeroWidthAnchors";

type ResolutionMode = "accept" | "reject";

/**
 * Whether the cell a paragraph's words would run into is itself pending: its
 * row or the cell (through nested tables' first cells) is inserted, deleted
 * or merged, so resolving may not keep it. A deletion is not pending for a
 * rejection of every change. Pending rows or cells elsewhere in the table leave
 * that cell where it is.
 */
export const runInCellPending = (
  table: PMNode,
  everyChange?: ResolutionMode,
): boolean => {
  // Rejecting everything keeps what was deleted, as if it had never been.
  const deletionStays = everyChange === "reject";
  let node: PMNode | null = table;
  while (node?.type.spec["tableRole"] === "table") {
    const row: PMNode | null = node.firstChild;
    if (!row) return false;
    if (row.attrs["trIns"] != null) return true;
    if (row.attrs["trDel"] != null && !deletionStays) return true;
    const cell: PMNode | null = row.firstChild;
    if (!cell) return false;
    const marker = cell.attrs["cellMarker"] as { kind: string } | null;
    if (marker != null && !(deletionStays && marker.kind === "del"))
      return true;
    node = cell.firstChild;
  }
  return false;
};

/**
 * Whether resolving every change takes this row out with what it records. The
 * table it stands in goes when all its rows do.
 */
const rowGoes = (row: PMNode, mode: ResolutionMode): boolean =>
  row.attrs[mode === "accept" ? "trDel" : "trIns"] != null;

/** Whether resolving every change in `mode` removes the whole table. */
export const tableGoes = (table: PMNode, mode: ResolutionMode): boolean => {
  if (table.type.spec["tableRole"] !== "table" || table.childCount === 0)
    return false;
  for (let index = 0; index < table.childCount; index++) {
    if (!rowGoes(table.child(index), mode)) return false;
  }
  return true;
};

/** The first paragraph of a table's first cell, descending into nested tables. */
const runInTarget = (
  doc: PMNode,
  tablePos: number,
  everyChange: ResolutionMode | undefined,
): { pos: number; node: PMNode } | null => {
  let node = doc.nodeAt(tablePos);
  let pos = tablePos;
  const table = node;
  const firstRow = table?.firstChild;
  if (!table || table.type.spec["tableRole"] !== "table" || !firstRow)
    return null;
  // A pending row or cell may not keep the cell the words would run into:
  // the paragraph keeps its own place instead.
  if (runInCellPending(table, everyChange)) return null;
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
  /**
   * Set when every change of the story resolves this way, not only some: what
   * it removes does not count as words left, and a rejection takes a deleted
   * row as staying.
   */
  everyChange?: ResolutionMode | undefined;
};

const paragraphAfterContentRemoval = (
  paragraph: PMNode,
  everyChange: ResolutionMode | undefined,
) =>
  everyChange
    ? paragraph.copy(
        Fragment.fromArray(
          paragraph.children.filter(
            (child) =>
              !child.marks.some(
                (mark) =>
                  REVISION_MARK_NAMES.has(mark.type.name) &&
                  trackedRevisionPathOf(mark).some((layer) =>
                    revisionLayerRemovesContent(layer, everyChange),
                  ),
              ),
          ),
        ),
      )
    : paragraph;

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
  everyChange,
}: RunParagraphIntoTableOptions): {
  position: number;
  map: StepMap;
  transfer: ParagraphPropertySourceTransfer;
} | null => {
  const paragraph = tr.doc.nodeAt(paragraphPos);
  if (paragraph?.type.name !== "paragraph") return null;
  if (expectParagraphAttrs(paragraph)._sectionProperties !== undefined)
    return null;
  const left = paragraphAfterContentRemoval(paragraph, everyChange);
  const remains = paragraphRemainsOf(left);
  if (
    !paragraphContentOf(left) &&
    (remains.blocks.length > 0 || remains.comments.length === 0)
  )
    return null;
  const tablePos = paragraphPos + paragraph.nodeSize;
  const target = runInTarget(tr.doc, tablePos, everyChange);
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
  const map = new StepMap([
    paragraphPos,
    0,
    levels - 1,
    closing,
    target.pos + 1 - closing,
    0,
  ]);
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
const markGoes = (paragraph: PMNode, mode: ResolutionMode): boolean => {
  const mark = expectParagraphAttrs(paragraph).pPrMark;
  if (!mark) return false;
  const added = mark.kind === "ins" || mark.kind === "moveTo";
  return added !== (mode === "accept");
};

/**
 * Whether resolving every change in `mode` runs `paragraph` on into `next`:
 * its mark goes and `next` is a table.
 */
type RunsIntoFollowingTableOptions = {
  parent: PMNode;
  index: number;
  mode: ResolutionMode;
};

/** The comments whose marks the inline content of `paragraph` carries; point comments mark nothing. */
const markedCommentIds = (paragraph: PMNode): Set<number> => {
  const ids = new Set<number>();
  paragraph.descendants((node) => {
    if (!node.isInline || node.type.name === COMMENT_REFERENCE_NODE_NAME)
      return true;
    for (const mark of node.marks) {
      const commentId: unknown = mark.attrs["commentId"];
      if (
        mark.type.name === "comment" &&
        mark.attrs["point"] !== true &&
        typeof commentId === "number"
      )
        ids.add(commentId);
    }
    return true;
  });
  return ids;
};

/**
 * `left`, what resolving leaves of `paragraph`, without the comment references
 * the resolution orphans: those of comments whose marked content here it
 * removes, all of it. A comment's reference follows what the comment still
 * covers, or goes with the comment, so it keeps nothing of the paragraph, as
 * removing the same content directly keeps nothing.
 */
const withoutOrphanedCommentReferences = (
  paragraph: PMNode,
  left: PMNode,
): PMNode => {
  const kept = markedCommentIds(left);
  const orphaned = new Set(
    [...markedCommentIds(paragraph)].filter((id) => !kept.has(id)),
  );
  if (orphaned.size === 0) return left;
  return left.copy(
    Fragment.fromArray(
      left.children.filter(
        (child) =>
          child.type.name !== COMMENT_REFERENCE_NODE_NAME ||
          !orphaned.has(expectCommentReferenceAttrs(child).commentId),
      ),
    ),
  );
};

/**
 * Whether resolving every change in `mode` leaves nothing of this paragraph:
 * its mark goes, and so does everything in it, the references of comments it
 * takes all the marked content of included. This is the one reading of a
 * block removed whole that pending joins, gaps and deletions share.
 */
export const paragraphGoes = (
  paragraph: PMNode,
  mode: ResolutionMode,
): boolean => {
  if (
    paragraph.type.name !== "paragraph" ||
    !markGoes(paragraph, mode) ||
    expectParagraphAttrs(paragraph)._sectionProperties != null
  )
    return false;
  const left = withoutOrphanedCommentReferences(
    paragraph,
    paragraphAfterContentRemoval(paragraph, mode),
  );
  if (paragraphContentOf(left)) return false;
  const remains = paragraphRemainsOf(left);
  return (
    remains.blocks.length === 0 &&
    remains.comments.length === 0 &&
    remains.inline.length === 0
  );
};

/**
 * Whether what stands right after a paragraph whose mark goes is itself removed
 * by this resolution, so the paragraph runs on into what stands after that, as
 * it would had the removed block never been there: a table it removes, or, when
 * rejecting, an inserted paragraph it empties with a table after it. (Accepted,
 * a deleted paragraph did stand there, and the join ended at it. Paragraphs join
 * among themselves without help.)
 */
export const paragraphBeforeRemovedBlock = ({
  parent,
  index,
  mode,
}: RunsIntoFollowingTableOptions): boolean => {
  const paragraph = parent.child(index);
  const next = parent.maybeChild(index + 1);
  if (
    paragraph.type.name !== "paragraph" ||
    !next ||
    !markGoes(paragraph, mode)
  ) {
    return false;
  }
  if (tableGoes(next, mode)) return true;
  if (mode !== "reject" || paragraphGoes(paragraph, mode)) return false;
  let following = index + 1;
  while (following < parent.childCount) {
    const block = parent.child(following);
    if (!paragraphGoes(block, mode)) break;
    following += 1;
  }
  return (
    following > index + 1 &&
    parent.maybeChild(following)?.type.spec["tableRole"] === "table"
  );
};

export const runsIntoFollowingTable = ({
  parent,
  index,
  mode,
}: RunsIntoFollowingTableOptions): boolean => {
  const paragraph = parent.child(index);
  const next = parent.maybeChild(index + 1);
  if (
    paragraph.type.name !== "paragraph" ||
    next?.type.spec["tableRole"] !== "table" ||
    !markGoes(paragraph, mode)
  ) {
    return false;
  }
  const remains = paragraphRemainsOf(paragraph);
  const holdsContent = paragraphContentOf(paragraph);
  const needsCommentHost =
    !holdsContent && remains.blocks.length === 0 && remains.comments.length > 0;
  let hasPrecedingCommentHost = false;
  for (
    let previousIndex = needsCommentHost ? index - 1 : -1;
    previousIndex >= 0;
    previousIndex -= 1
  ) {
    const previous = parent.child(previousIndex);
    if (previous.type.name !== "paragraph") break;
    if (
      !markGoes(previous, mode) ||
      expectParagraphAttrs(previous)._sectionProperties !== undefined
    ) {
      hasPrecedingCommentHost = true;
      break;
    }
    const retained = paragraphAfterContentRemoval(previous, mode);
    if (paragraphContentOf(retained)) break;
    if (paragraphRemainsOf(retained).blocks.length > 0) {
      hasPrecedingCommentHost = true;
      break;
    }
  }
  return (
    paragraphMarkOutcome({
      follows: "table",
      holdsContent,
      markWasAdded:
        expectParagraphAttrs(paragraph).pPrMark?.kind === "ins" ||
        expectParagraphAttrs(paragraph).pPrMark?.kind === "moveTo",
      containerStillEndsWithParagraph: false,
      followingTableCanReceiveRunIn:
        expectParagraphAttrs(paragraph)._sectionProperties === undefined &&
        !runInCellPending(next, mode),
      needsFollowingCommentHost: needsCommentHost && !hasPrecedingCommentHost,
    }).kind === "run-in"
  );
};

/**
 * Stand each block this resolution removes before the paragraph whose mark goes
 * right before it, so that paragraph meets what it will run on into. The block
 * is going, so where it goes from does not show; the paragraph's words keep
 * mapping to themselves.
 */
const moveRemovedBlocksBeforeParagraphs = (
  tr: Transform,
  mode: ResolutionMode,
): StepMap[] => {
  const maps: StepMap[] = [];
  for (;;) {
    let move: { position: number; paragraph: PMNode; block: PMNode } | null =
      null;
    tr.doc.descendants((node, position, parent, index) => {
      if (move !== null) return false;
      if (node.type.name !== "paragraph") return !node.isTextblock;
      const block = parent?.maybeChild(index + 1);
      if (
        parent &&
        block &&
        paragraphBeforeRemovedBlock({ parent, index, mode })
      ) {
        move = { position, paragraph: node, block };
      }
      return false;
    });
    if (move === null) return maps;
    const { position, paragraph, block } = move as {
      position: number;
      paragraph: PMNode;
      block: PMNode;
    };
    const blockPosition = position + paragraph.nodeSize;
    tr.step(
      new ReplaceStep(
        position,
        blockPosition + block.nodeSize,
        new Slice(Fragment.fromArray([block, paragraph]), 0, 0),
      ),
    );
    maps.push(
      new StepMap([
        position,
        0,
        block.nodeSize,
        blockPosition,
        block.nodeSize,
        0,
      ]),
    );
  }
};

/**
 * Before every change in a story is resolved: run each paragraph whose mark
 * goes and that sits right before a table on into that table. Returns each
 * paragraph that took words (its position after the step count given) and
 * the paragraph-property sources handed over.
 */
export const runParagraphsIntoTables = (
  tr: Transform,
  mode: ResolutionMode,
  styleResolver: RunStyleResolver | null,
): {
  targets: { position: number; step: number }[];
  maps: StepMap[];
  transfers: ParagraphPropertySourceTransfer[];
} => {
  const tried = new Set<number>();
  const targets: { position: number; step: number }[] = [];
  const transfers: ParagraphPropertySourceTransfer[] = [];
  const maps: StepMap[] = moveRemovedBlocksBeforeParagraphs(tr, mode);
  for (;;) {
    let candidate: number | null = null;
    tr.doc.descendants((node, pos, parent, index) => {
      if (node.type.name !== "paragraph") return !node.isTextblock;
      if (
        parent &&
        runsIntoFollowingTable({
          parent,
          index,
          mode,
        }) &&
        !tried.has(pos)
      ) {
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
      everyChange: mode,
    });
    if (ranIn !== null) {
      targets.push({ position: ranIn.position, step: tr.steps.length });
      transfers.push(ranIn.transfer);
      maps.push(ranIn.map);
      tried.clear();
    }
  }
};
