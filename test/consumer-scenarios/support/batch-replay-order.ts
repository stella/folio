import {
  BATCH_ANNOTATION_KINDS,
  type BatchAnnotationKind,
  type BatchOrderAddress,
  contentRank,
  deletionRemovesBlock,
  type FinalBreakCandidate,
  type FinalBreakCarrier,
  finalBreakCarrier,
  groupInsertionRuns,
  orderAddressingBeforeRemoval,
  orderBatchOperations,
} from "@stll/docx-core/ops";

import type { AppliedOperation, Mode } from "./operations.ts";
import type { Row } from "./oracle.ts";

type AnyOperation = AppliedOperation;
type Handle = { blockId: string; startOffset: number; endOffset: number };

const INSERTIONS = new Set([
  "insertAfterBlock",
  "insertBeforeBlock",
  "insertTable",
]);
const ANNOTATION_KIND_OF = new Map<string, BatchAnnotationKind>([
  ["formatRange", BATCH_ANNOTATION_KINDS.FORMATTING],
  ["commentOnRange", BATCH_ANNOTATION_KINDS.COMMENT],
  ["commentOnBlock", BATCH_ANNOTATION_KINDS.COMMENT],
]);
const RANGED = new Set(["replaceRange", "formatRange", "commentOnRange"]);

const handleOf = (operation: AnyOperation): Handle | null =>
  RANGED.has(operation.type) ? (operation["range"] as Handle) : null;
const blockIdOf = (operation: AnyOperation): string =>
  handleOf(operation)?.blockId ?? String(operation["blockId"]);

/** An offset past every one inside a block (`placement` keeps them below 10^8). */
const AFTER_EVERY_OFFSET = 99_999_999;

/** Whether an insertion goes before its anchor rather than after it. */
const insertsBefore = (operation: AnyOperation): boolean =>
  operation.type === "insertBeforeBlock" ||
  (operation.type === "insertTable" && operation["position"] === "before");

/**
 * The gap among `preRows` that an insertion anchored on a block they lack
 * lands in: a block accepting removes, or a cell of a row or table it
 * removes, named from `liveRows`. Beside a cell it goes beside the outermost
 * table, at its first or last row that stays; with none staying, and beside
 * a removed block outside tables, it goes before the next row that stays,
 * the one gap the blocks around the removed ones share.
 */
const absentAnchorGap = (
  operation: AnyOperation,
  preRows: readonly Row[],
  liveRows: readonly Row[],
): number | undefined => {
  const id = blockIdOf(operation);
  if (!INSERTIONS.has(operation.type) || preRows.some((row) => row.id === id))
    return undefined;
  const at = liveRows.findIndex((row) => row.id === id);
  if (at < 0) return undefined;
  const outer = liveRows[at]?.table?.outerTableIndex;
  const span =
    outer === undefined
      ? [at]
      : liveRows.flatMap((row, index) =>
          row.table?.outerTableIndex === outer ? [index] : [],
        );
  const preIndex = new Map(preRows.map((row, index) => [row.id, index]));
  const staying = span.flatMap((index) => {
    const kept = preIndex.get(liveRows[index]?.id ?? "");
    return kept === undefined ? [] : [kept];
  });
  const first = staying.at(0);
  const last = staying.at(-1);
  if (first !== undefined && last !== undefined)
    return insertsBefore(operation) ? first : last + 1;
  for (const row of liveRows.slice((span.at(-1) ?? at) + 1)) {
    const kept = preIndex.get(row.id);
    if (kept !== undefined) return kept;
  }
  return preRows.length;
};

/**
 * The paragraphs among `rows`, outside tables, that accepting removes whole:
 * a reader lists one as a blank block, and its own mark is pending deletion
 * (`markDeleted`). The two sides of such paragraphs are one gap, as the
 * engine places blocks, unless words before them run on through them: the
 * paragraph before the blank ones has a mark pending deletion too, and the
 * join stops the gap. A cell paragraph stays: the gap is the table's.
 */
export const blanksGoingWhole = (
  rows: readonly Row[],
  markDeleted: ReadonlySet<string>,
): Set<string> => {
  const goWhole = new Set<string>();
  let joinRunsOn = false;
  for (const { id, text, table } of rows) {
    if (table !== undefined) {
      joinRunsOn = false;
      continue;
    }
    if (text === "" && markDeleted.has(id)) {
      if (!joinRunsOn) goWhole.add(id);
      continue;
    }
    joinRunsOn = markDeleted.has(id);
  }
  return goWhole;
};

const placement = (
  operation: AnyOperation,
  preRows: readonly Row[],
  joinsInto: ReadonlyMap<string, PendingJoinTarget> = new Map(),
  liveRows: readonly Row[] = preRows,
): [number, number, number] => {
  const index = preRows.findIndex((row) => row.id === blockIdOf(operation));
  const text = preRows[index]?.text ?? "";
  const handle = handleOf(operation);
  if (handle) return [index, 1, handle.startOffset];
  const absentGap = absentAnchorGap(operation, preRows, liveRows);
  if (absentGap !== undefined) return [absentGap, 0, 0];
  // A block added beside a table cell, a paragraph or a whole table, goes
  // outside the table: before its first row or after its last, so every
  // anchor in one outermost table shares a gap.
  const outer = preRows[index]?.table?.outerTableIndex;
  if (outer !== undefined && INSERTIONS.has(operation.type)) {
    const rows = preRows.flatMap((row, at) =>
      row.table?.outerTableIndex === outer ? [at] : [],
    );
    return insertsBefore(operation)
      ? [rows.at(0) ?? index, 0, 0]
      : [(rows.at(-1) ?? index) + 1, 0, 0];
  }
  // A block put after a paragraph whose pending join runs into a later one
  // lands after that one as the batch read it, and runs there: ahead of the
  // edits inside the joined paragraphs. Into one with no words left, it
  // lands before that one, still as the batch read it: it runs ahead of the
  // edits inside it, which alone would give the join words to run into and
  // put the block after it. A paragraph whose words are all pending deletion,
  // with none running on through it, joins nothing into one that keeps
  // words: a block put after it lands right after it, as one put after any
  // paragraph does. An empty one reads as one paragraph with the one its join
  // leaves, so the block lands after that one.
  const joined = joinsInto.get(blockIdOf(operation));
  const after =
    operation.type === "insertAfterBlock" ||
    (operation.type === "insertTable" && operation["position"] !== "before");
  const joinsWords =
    joined !== undefined &&
    !(joined.keepsWords && joined.source === PENDING_JOIN_SOURCES.GOES_WHOLE);
  if (after && joined !== undefined && joinsWords) {
    const survivor = preRows.findIndex((row) => row.id === joined.survivor);
    if (survivor >= 0)
      return joined.keepsWords
        ? [survivor + 1, 0, 0]
        : [survivor, 1, AFTER_EVERY_OFFSET];
  }
  // A block put before that paragraph with no words left lands in the same
  // gap, and runs there too: alone after an edit gave it words, it would go
  // before the paragraphs the join takes in.
  const before =
    operation.type === "insertBeforeBlock" ||
    (operation.type === "insertTable" && operation["position"] === "before");
  if (
    before &&
    [...joinsInto.values()].some(
      ({ survivor, keepsWords }) =>
        !keepsWords && survivor === blockIdOf(operation),
    )
  )
    return [index, 1, AFTER_EVERY_OFFSET];
  // A block put before a paragraph of a join into one that keeps words goes
  // before the first paragraph the join takes in (`pendingJoinParagraphs`),
  // the gap a block put before that one lands in.
  if (before && (text !== "" || joined?.keepsWords === true)) {
    let first = index;
    for (let at = index - 1; at >= 0; at -= 1) {
      const row = preRows[at];
      if (row === undefined) break;
      const source = joinsInto.get(row.id);
      if (source === undefined || !sameContainerRows(row, preRows[index]))
        break;
      if (source.source === PENDING_JOIN_SOURCES.JOINS_IN) first = at;
    }
    if (first !== index) return [first, 0, 0];
  }
  switch (operation.type) {
    case "replaceInBlock":
      return [index, 1, Math.max(0, text.indexOf(String(operation["find"])))];
    case "splitBlock":
      return [index, 1, Number(operation["offset"])];
    case "mergeBlockWithNext":
    case "insertAfterBlock":
    case "insertTable":
      return [index + 1, 0, 0];
    case "setBlockParagraphProperties":
    case "insertBeforeBlock":
      return [index, 0, 0];
    default:
      return [index, 1, 0];
  }
};

const sameContainerRows = (left: Row, right: Row | undefined): right is Row =>
  right !== undefined &&
  `${left.table?.tableIndex}:${left.table?.rowIndex}:${left.table?.cellIndex}` ===
    `${right.table?.tableIndex}:${right.table?.rowIndex}:${right.table?.cellIndex}`;

/**
 * What accepting does with a paragraph of a pending join, as the engine reads
 * it (`pendingJoinParagraphs`): its words, or its empty content, run into the
 * paragraph the join leaves (`joinsIn`); its words are all pending deletion,
 * but words before it run on through it (`runsOn`); or, its words all pending
 * deletion with none running on, it goes whole (`goesWhole`).
 */
export const PENDING_JOIN_SOURCES = {
  JOINS_IN: "joinsIn",
  RUNS_ON: "runsOn",
  GOES_WHOLE: "goesWhole",
} as const;

export type PendingJoinSource =
  (typeof PENDING_JOIN_SOURCES)[keyof typeof PENDING_JOIN_SOURCES];

/**
 * The paragraph a pending join runs into, whether it keeps words, and what
 * accepting does with the joining paragraph.
 */
export type PendingJoinTarget = {
  survivor: string;
  keepsWords: boolean;
  source: PendingJoinSource;
};

type PendingJoinTargetsOptions = {
  rows: readonly Row[];
  /** The paragraphs whose mark is pending deletion (`markDeletedIds`). */
  markDeleted: ReadonlySet<string>;
  /** The paragraphs holding words pending deletion (`wordsDeletedIds`). */
  wordsDeleted: ReadonlySet<string>;
};

/**
 * The paragraph each pending join in `rows` runs into: from a paragraph whose
 * mark is pending deletion (`markDeleted`) on past the next ones whose marks
 * go too, to the first one whose mark stays, in the same container. The
 * engine places a block put after any paragraph of the join after that one
 * when it keeps words, and in the gap before it otherwise. A row with no words
 * left is empty unless it holds words pending deletion.
 */
export const pendingJoinTargets = ({
  rows,
  markDeleted,
  wordsDeleted,
}: PendingJoinTargetsOptions): Map<string, PendingJoinTarget> => {
  const targets = new Map<string, PendingJoinTarget>();
  let wordsRunOn = false;
  for (const [index, row] of rows.entries()) {
    if (!markDeleted.has(row.id)) {
      wordsRunOn = false;
      continue;
    }
    const previous = rows[index - 1];
    if (previous === undefined || !sameContainerRows(row, previous))
      wordsRunOn = false;
    let source: PendingJoinSource = PENDING_JOIN_SOURCES.JOINS_IN;
    if (row.text === "" && wordsDeleted.has(row.id))
      source = wordsRunOn
        ? PENDING_JOIN_SOURCES.RUNS_ON
        : PENDING_JOIN_SOURCES.GOES_WHOLE;
    if (row.text !== "") wordsRunOn = true;
    let at = index + 1;
    while (
      markDeleted.has(rows[at]?.id ?? "") &&
      sameContainerRows(row, rows[at])
    )
      at += 1;
    const survivor = rows[at];
    if (sameContainerRows(row, survivor))
      targets.set(row.id, {
        survivor: survivor.id,
        keepsWords: survivor.text.length > 0,
        source,
      });
  }
  return targets;
};

type SequentialGroupsOptions = {
  applied: readonly AnyOperation[];
  /** The rows the batch's outcome is read against (accepted, for tracked changes). */
  preRows: readonly Row[];
  /**
   * The live rows the operations name, when `preRows` are not those. A
   * column edit may anchor in a cell of a row pending deletion, which only
   * the live table still has; column edits order among themselves, so they
   * all take live coordinates.
   */
  liveRows?: readonly Row[];
  mode: Mode;
  /** The paragraph each pending join runs into (`pendingJoinTargets`). */
  joinsInto?: ReadonlyMap<string, PendingJoinTarget>;
  /** Blank paragraphs of `preRows` accepting removes whole (`blanksGoingWhole`). */
  goWhole?: ReadonlySet<string>;
};

export const sequentialGroups = ({
  applied,
  preRows,
  liveRows = preRows,
  mode,
  joinsInto = new Map(),
  goWhole = new Set(),
}: SequentialGroupsOptions): AnyOperation[][] => {
  /**
   * Where the gap at `at` starts: before the blank paragraphs right before
   * it that accepting removes whole.
   */
  const gapStart = (at: number): number => {
    let start = at;
    while (start > 0 && goWhole.has(preRows[start - 1]?.id ?? "")) start -= 1;
    return start;
  };
  /**
   * The gap an insertion lands in, which decides the run it replays in. The
   * two sides of blank paragraphs accepting removes whole are one gap: a
   * block put before the paragraph after them lands before them, beside one
   * put after the paragraph before them, in the order the batch names them.
   * One put after such a blank paragraph runs with them too: one at a time
   * ahead of a block put before the paragraphs after it, it would stop that
   * block's walk back across the blanks.
   */
  const insertionGapKey = (operation: AnyOperation): number => {
    const [block, inside, offset] = placement(
      operation,
      preRows,
      joinsInto,
      liveRows,
    );
    if (inside !== 0)
      return block * 1_000_000_000 + inside * 100_000_000 + offset;
    const id = blockIdOf(operation);
    const from =
      goWhole.has(id) && !insertsBefore(operation)
        ? preRows.findIndex((row) => row.id === id)
        : block;
    return gapStart(from) * 1_000_000_000 + offset;
  };
  const indexed = applied.map((operation, index) => ({ operation, index }));
  const deleted = new Set(
    applied
      .filter((operation) => operation.type === "deleteBlock")
      .map(blockIdOf),
  );
  const containers = new Map<string, Row[]>();
  for (const row of preRows) {
    const key = row.table
      ? `${row.table.tableIndex}:${row.table.rowIndex}:${row.table.cellIndex}`
      : "story";
    const rows = containers.get(key) ?? [];
    rows.push(row);
    containers.set(key, rows);
  }
  /**
   * The paragraph that settles a deleted container-final paragraph, as the
   * engine reads it (`finalBreakCarrier`): the paragraphs before it in its
   * container, up to a table (rows of another container, `preRows` being in
   * document order) or the container's start.
   */
  const breakCarrierOf = (last: Row): FinalBreakCarrier<Row> => {
    const before: FinalBreakCandidate<Row>[] = [];
    for (let at = preRows.indexOf(last) - 1; at >= 0; at -= 1) {
      const row = preRows.at(at);
      if (!sameContainerRows(last, row)) break;
      before.push({ paragraph: row, deleted: deleted.has(row.id) });
    }
    return finalBreakCarrier(before);
  };
  /**
   * Whether the batch places a block in the gap before a deleted final
   * paragraph's run: the paragraph and the ones before it in its container the
   * batch deletes too. Whether the paragraph goes or stays emptied is read
   * from what precedes it once the batch is finished, in every mode, so its
   * deletion is replayed after such a block lands.
   */
  const placesBeforeRun = (last: Row): boolean => {
    const lastAt = preRows.indexOf(last);
    let start = lastAt;
    while (start > 0) {
      const row = preRows.at(start - 1);
      if (!sameContainerRows(last, row) || !deleted.has(row.id)) break;
      start -= 1;
    }
    return applied.some((operation) => {
      if (!INSERTIONS.has(operation.type)) return false;
      const anchor = preRows.find((row) => row.id === blockIdOf(operation));
      if (anchor === undefined) return false;
      const [index] = placement(operation, preRows, joinsInto, liveRows);
      if (index < start || index > lastAt) return false;
      // A paragraph placed beside a cell goes outside the table, into the
      // container the table stands in.
      return anchor.table !== undefined && last.table === undefined
        ? operation.type !== "insertTable"
        : sameContainerRows(last, anchor);
    });
  };
  const retiringFinals = new Set<string>();
  const finalBlocks = new Set<string>();
  const uncarriedFinals = new Set<string>();
  /** A late final deletion's block, mapped to its run's first paragraph, deleted after it. */
  const runStartAfter = new Map<string, string>();
  // A suggestion reads the live rows, where a paragraph pending deletion is a
  // blank block: the block before such blocks ends its container as the batch
  // sees it.
  const finalBeforeBlanks = new Set<string>();
  for (const [container, rows] of containers) {
    const last = rows.at(-1);
    if (!last || (container === "story" && last !== preRows.at(-1))) continue;
    finalBlocks.add(last.id);
    const beforeBlanks = rows.findLast(({ text }) => text.length > 0);
    if (mode === "suggested" && beforeBlanks) {
      finalBeforeBlanks.add(beforeBlanks.id);
    }
    const appended = applied.some(
      (operation) =>
        operation.type === "insertAfterBlock" &&
        blockIdOf(operation) === last.id,
    );
    if (appended || !deleted.has(last.id)) continue;
    const carrier = breakCarrierOf(last);
    if (carrier.type !== "kept") uncarriedFinals.add(last.id);
    // A sole paragraph with one inserted before it is deleted first, and the
    // insertion takes its place. A paragraph that ends a deleted run stays until
    // the run is gone. Applied directly only such a replacement waits, since
    // nothing before it holds a break to retire.
    const replaced = applied.some(
      (operation) =>
        operation.type === "insertBeforeBlock" &&
        blockIdOf(operation) === last.id,
    );
    const lastAt = preRows.indexOf(last);
    const endsARun =
      lastAt > 0 && sameContainerRows(last, preRows.at(lastAt - 1));
    const runsLate =
      mode === "tracked-changes"
        ? !replaced || endsARun
        : mode === "direct" && carrier.type !== "kept" && replaced && endsARun;
    const placesInRun = placesBeforeRun(last);
    if (!runsLate && !placesInRun) continue;
    retiringFinals.add(last.id);
    // A run the batch deletes back to its container's start or a table
    // collapses into its first paragraph. Run late, the final paragraph's
    // deletion has to find that paragraph still standing to take its
    // properties, so the first paragraph's own deletion follows it. A block
    // the batch places in the run's gap takes the break instead.
    if (carrier.type === "runStart" && !placesInRun) {
      runStartAfter.set(last.id, carrier.paragraph.id);
    }
  }
  const isColumnEdit = (operation: AnyOperation): boolean =>
    operation.type === "insertTableColumn" ||
    operation.type === "deleteTableColumn";
  const deletedSplitTargets = new Set(
    applied
      .filter(
        (operation) =>
          mode === "tracked-changes" &&
          operation.type === "deleteBlock" &&
          applied.some(
            (candidate) =>
              candidate.type === "insertTable" &&
              candidate["position"] !== "before" &&
              blockIdOf(candidate) === blockIdOf(operation),
          ),
      )
      .map(blockIdOf),
  );
  const columnLocation = (operation: AnyOperation): [number, number] => {
    const targetId = blockIdOf(operation);
    const target = liveRows.find((row) => row.id === targetId);
    if (!target?.table)
      throw new Error(
        `Applied ${operation.type} anchor ${targetId} has no source table.`,
      );
    const sourceColumn = target.table.gridColumnIndex;
    const column =
      operation.type === "insertTableColumn" &&
      operation["position"] !== "before"
        ? sourceColumn + target.table.columnSpan
        : sourceColumn;
    return [target.table.tableIndex, column];
  };
  const ordered = orderBatchOperations(
    indexed.filter(({ operation }) => !ANNOTATION_KIND_OF.has(operation.type)),
    ({ operation, index }): BatchOrderAddress => {
      if (isColumnEdit(operation)) {
        const [tablePosition, columnIndex] = columnLocation(operation);
        return {
          type: "column",
          tablePosition,
          columnIndex,
          insertion: operation.type === "insertTableColumn",
          inputIndex: index,
        };
      }
      const [block, inside, offset] = placement(
        operation,
        preRows,
        joinsInto,
        liveRows,
      );
      const hasDeletedSplit = deletedSplitTargets.has(blockIdOf(operation));
      const position =
        hasDeletedSplit && operation.type === "deleteBlock"
          ? (block + 1) * 1_000_000_000
          : block * 1_000_000_000 + inside * 100_000_000 + offset;
      return {
        type: "content",
        position,
        rank: contentRank(operation, { hasDeletedSplit }),
        inputIndex: index,
      };
    },
  );
  const rest = ordered.filter(({ operation }) => !isColumnEdit(operation));
  const columns = ordered.filter(({ operation }) => isColumnEdit(operation));
  const inputOrder = new Map(applied.map(({ id }, index) => [id, index]));
  const groups = groupInsertionRuns(
    rest.map(({ operation }) => operation),
    (operation) =>
      INSERTIONS.has(operation.type) ? insertionGapKey(operation) : null,
    (operation) => operation.type === "setBlockParagraphProperties",
  ).map((group) =>
    group.toSorted(
      (left, right) =>
        (inputOrder.get(left.id) ?? Number.MAX_SAFE_INTEGER) -
        (inputOrder.get(right.id) ?? Number.MAX_SAFE_INTEGER),
    ),
  );
  // A deletion removes or retires its block, so an insertion before it can
  // only be stated one at a time while the block still stands. The batch
  // keeps that gap, so the oracle runs the whole run ahead of the deletion.
  // A container's final paragraph differs. With no surviving paragraph to
  // retire its mark onto, a paragraph inserted first would carry the break,
  // or take the deleted paragraph's place directly, and the batch deletes
  // first. In a suggestion the batch hands the break to a paragraph inserted
  // into the gap before it, whichever side of the gap the insertion is stated
  // from.
  const gapDeletion = (operation: AnyOperation): string | undefined => {
    const id = blockIdOf(operation);
    if (operation.type === "insertBeforeBlock") {
      const retires = mode !== "suggested" && uncarriedFinals.has(id);
      return deleted.has(id) && !retires ? id : undefined;
    }
    if (operation.type !== "insertAfterBlock" || mode === "direct") {
      return undefined;
    }
    const index = preRows.findIndex((row) => row.id === id);
    const next = preRows.at(index + 1);
    return next !== undefined &&
      deleted.has(next.id) &&
      (finalBlocks.has(next.id) || finalBeforeBlanks.has(next.id)) &&
      sameContainerRows(next, preRows.at(index))
      ? next.id
      : undefined;
  };
  for (const { operation } of rest) {
    const deletedId = gapDeletion(operation);
    if (deletedId === undefined) continue;
    const from = groups.findIndex((group) => group.includes(operation));
    const target = groups.findIndex(
      (group) =>
        group.length === 1 &&
        group[0]?.type === "deleteBlock" &&
        blockIdOf(group[0]) === deletedId,
    );
    // The deletion moves, not the run: what else is at the gap keeps its place
    // beside the run, in the order the batch runs them.
    const deletion = groups[target];
    if (deletion && target >= 0 && target < from) {
      groups.splice(target, 1);
      groups.splice(from, 0, deletion);
    }
  }
  // The shared batch order: a run addressing a block goes before the run that
  // deletes or merges it away, as the engine executes the batch.
  const appendedAfter = (id: string) =>
    applied.some(
      (operation) =>
        blockIdOf(operation) === id &&
        (operation.type === "insertAfterBlock" ||
          (operation.type === "insertTable" &&
            operation["position"] !== "before")),
    );
  const removedBy = (operation: AnyOperation): string[] => {
    const id = blockIdOf(operation);
    if (operation.type === "deleteBlock") {
      const row = preRows.find((candidate) => candidate.id === id);
      return row === undefined ||
        deletionRemovesBlock({
          endsContainer: finalBlocks.has(id) && !appendedAfter(id),
          carrier: breakCarrierOf(row).type,
        })
        ? [id]
        : [];
    }
    // Removing a table row or table takes its anchor with it.
    if (operation.type === "deleteTableRow" || operation.type === "deleteTable")
      return [id];
    if (operation.type !== "mergeBlockWithNext") return [];
    const index = preRows.findIndex((row) => row.id === id);
    const row = preRows.at(index);
    const next = preRows.at(index + 1);
    if (row === undefined || !sameContainerRows(row, next)) return [];
    // Applied directly, the join retires one paragraph's id and rewrites the
    // other; tracked, both stay until the merge is accepted.
    return mode === "direct" ? [row.id, next.id] : [next.id];
  };
  // A paragraph inserted before one a direct merge joins away goes into the
  // gap the join keeps, after the join, as the engine runs it.
  const directlyJoinedAway = new Set(
    mode === "direct"
      ? applied
          .filter(({ type }) => type === "mergeBlockWithNext")
          .map(blockIdOf)
      : [],
  );
  const reordered = orderAddressingBeforeRemoval(groups, (group) => ({
    addresses: group.flatMap((operation) =>
      operation.type === "deleteBlock" ||
      operation.type === "deleteTableRow" ||
      operation.type === "deleteTable" ||
      (operation.type === "insertBeforeBlock" &&
        directlyJoinedAway.has(blockIdOf(operation)))
        ? []
        : [blockIdOf(operation)],
    ),
    removes: group.flatMap(removedBy),
  }));
  const annotations = orderBatchOperations(
    indexed.flatMap(({ operation, index }) => {
      const kind = ANNOTATION_KIND_OF.get(operation.type);
      return kind === undefined ? [] : [{ operation, index, kind }];
    }),
    ({ operation, index, kind }): BatchOrderAddress => {
      const [position] = placement(operation, preRows);
      return { type: "annotation", kind, position, inputIndex: index };
    },
  );
  // The batch retires a final paragraph's break once, over the finished batch,
  // when what precedes it is settled. Until then the paragraphs before it can
  // still be written to and placed beside, so its deletion runs late. A merge
  // already reads that break as gone, so the deletion runs before any merge.
  const retiresBreak = (group: readonly AnyOperation[]): boolean =>
    group.every(
      (operation) =>
        operation.type === "deleteBlock" &&
        retiringFinals.has(blockIdOf(operation)),
    );
  // A run's first paragraph is deleted right after the final paragraph's
  // retirement that collapses the run into it.
  const retirements = reordered.filter(retiresBreak).flatMap((group) => {
    const head = group[0];
    const runStart =
      head === undefined ? undefined : runStartAfter.get(blockIdOf(head));
    const follower = reordered.find(
      (candidate) =>
        runStart !== undefined &&
        candidate.length === 1 &&
        candidate[0]?.type === "deleteBlock" &&
        blockIdOf(candidate[0]) === runStart,
    );
    return follower === undefined ? [group] : [group, follower];
  });
  const others = reordered.filter((group) => !retirements.includes(group));
  const firstMerge = others.findIndex((group) =>
    group.some(({ type }) => type === "mergeBlockWithNext"),
  );
  const retirementsAt = firstMerge < 0 ? others.length : firstMerge;
  return [
    ...annotations.map(({ operation }) => [operation]),
    ...others.slice(0, retirementsAt),
    ...retirements,
    ...others.slice(retirementsAt),
    ...columns.map(({ operation }) => [operation]),
  ];
};
