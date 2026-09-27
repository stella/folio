/**
 * Which operations of one batch may not share their targets.
 *
 * A batch resolves every operation against the document as it was read and
 * applies them from the end of the document backwards, so an operation never
 * moves the positions an earlier-placed one resolved to. That holds only while
 * the operations' targets are disjoint. Two operations on the same stretch of
 * text, or an operation inside a paragraph another one deletes, rewrites,
 * splits or merges away, are applied against coordinates the first one has
 * already invalidated — and the second then edits whatever now sits there,
 * which can be a paragraph neither operation named.
 *
 * Each operation therefore states what it claims before anything is applied,
 * and the later of two operations whose claims conflict is refused
 * (`overlappingOperation`). The caller re-reads the document after the batch
 * and sends the refused operation again against the result. What remains
 * applies exactly as the same operations applied one at a time, from the end
 * of the document backwards, would.
 *
 * Positions are the ones the batch resolved against. A block is keyed by the
 * position before its node.
 */

type PositionRange = { from: number; to: number };

/** The text an operation rewrites (`text`), only marks (`annotation`), or breaks (`split`). */
type InlineClaim = {
  type: "text" | "annotation" | "split";
  block: number;
  from: number;
  to: number;
};

export type BatchClaim =
  | InlineClaim
  /** Every character of the block: `replaceBlock` changing its text. */
  | { type: "rewriteBlock"; block: number }
  /** The block's paragraph properties, and nothing of its text. */
  | { type: "paragraphProperties"; block: number }
  | {
      type: "deleteBlock";
      block: number;
      /** Where the node ends. */
      end: number;
      /**
       * The paragraph survives, emptied: it ends its container. A block the
       * batch inserts at `end` lands first (the batch runs backwards), so the
       * paragraph no longer ends anything when the deletion runs and goes as
       * any other does — see `followedByInsertion`.
       */
      keepsParagraph: boolean;
      /** The node leaves the document while the batch is still applying. */
      removesNode: boolean;
      /** Applied directly rather than tracked. */
      direct: boolean;
    }
  /**
   * Joins `block` with the block starting at `next`, where `block` ends.
   * Applied directly the two nodes become one while the batch is still
   * applying (`joinsNow`); tracked, only `block`'s paragraph mark is marked
   * deleted and both paragraphs stay as they are until the change is accepted.
   */
  | { type: "merge"; block: number; next: number; joinsNow: boolean }
  /** A new block between two others, at `at`. */
  | { type: "insertion"; at: number }
  /**
   * A table, a row or a column's cells, removed. A removed row or column
   * leaves the table to the other structural operations of the batch; a
   * removed table leaves nothing to them.
   */
  | {
      type: "tableRemoval";
      table: number;
      wholeTable: boolean;
      axis: "row" | "column" | "table";
      ranges: readonly PositionRange[];
    }
  /** Any other change to a table's grid: rows or columns added, cells merged or split. */
  | { type: "tableStructure"; table: number; axis: "row" | "column" | "cell" }
  | { type: "unclaimed" };

type BlockRole =
  | { kind: InlineClaim["type"]; claim: InlineClaim }
  | { kind: "rewriteBlock" | "paragraphProperties" }
  | { kind: "deleteBlock"; keepsParagraph: boolean; removesNode: boolean; direct: boolean }
  /** The block a merge ends (`own`), or the one it pulls in (`next`). */
  | { kind: "mergeOwn" | "mergeNext"; joinsNow: boolean };

type BlockEntry = { operationId: string; role: BlockRole };

const blockRolesOf = (claim: BatchClaim): { block: number; role: BlockRole }[] => {
  switch (claim.type) {
    case "text":
    case "annotation":
    case "split":
      return [{ block: claim.block, role: { kind: claim.type, claim } }];
    case "rewriteBlock":
    case "paragraphProperties":
      return [{ block: claim.block, role: { kind: claim.type } }];
    case "deleteBlock":
      return [
        {
          block: claim.block,
          role: {
            kind: "deleteBlock",
            keepsParagraph: claim.keepsParagraph,
            removesNode: claim.removesNode,
            direct: claim.direct,
          },
        },
      ];
    case "merge":
      return [
        { block: claim.block, role: { kind: "mergeOwn", joinsNow: claim.joinsNow } },
        { block: claim.next, role: { kind: "mergeNext", joinsNow: claim.joinsNow } },
      ];
    default:
      return [];
  }
};

type DeleteBlockRole = Extract<BlockRole, { kind: "deleteBlock" }>;

/**
 * A deletion of the paragraph that ends its container, once the batch inserts
 * a block after it: the insertion lands first, the paragraph then ends
 * nothing, and it goes as any other — tracked, its mark is deleted into the
 * inserted block and its properties with it; directly, the node goes.
 */
const followedByInsertion = (role: DeleteBlockRole): DeleteBlockRole => ({
  ...role,
  keepsParagraph: false,
  removesNode: role.direct,
});

const isPoint = ({ from, to }: PositionRange): boolean => from === to;

/**
 * Two stretches of one block conflict when they overlap, or start at the same
 * position: applied from the end backwards, the one applied second would then
 * find the first one's text where its own range begins. Two insertions at one
 * point are the exception — each writes new text and keeps the other's order.
 */
const inlineRangesConflict = (left: PositionRange, right: PositionRange): boolean =>
  (left.from < right.to && right.from < left.to) ||
  (left.from === right.from && !(isPoint(left) && isPoint(right)));

const contains = (outer: PositionRange, inner: PositionRange): boolean =>
  outer.from <= inner.from && inner.to <= outer.to;

/** Whether two claims on the SAME block conflict; symmetric. */
/**
 * Which of two roles the rules below are written from: the one ranked
 * higher is put on the left, so each pair is decided in one place.
 */
const ROLE_RANK: Record<BlockRole["kind"], number> = {
  deleteBlock: 4,
  rewriteBlock: 3,
  mergeOwn: 2,
  mergeNext: 2,
  paragraphProperties: 1,
  split: 0,
  text: 0,
  annotation: 0,
};

const rolesConflict = (left: BlockRole, right: BlockRole): boolean => {
  if (ROLE_RANK[right.kind] > ROLE_RANK[left.kind]) {
    return rolesConflict(right, left);
  }
  switch (left.kind) {
    case "deleteBlock":
      switch (right.kind) {
        // An emptied paragraph that stays is still a paragraph to format.
        case "paragraphProperties":
          return !left.keepsParagraph;
        // A merge into the deleted block joins across its deleted mark while
        // tracked; applied directly, the block is gone and the merge would
        // join whatever follows it.
        case "mergeNext":
          return left.removesNode;
        default:
          return true;
      }
    case "rewriteBlock":
      switch (right.kind) {
        case "paragraphProperties":
        case "mergeNext":
          return false;
        // Tracked, the merge marks the paragraph's break and the rewrite its
        // words, and the two compose. Applied directly, the join lands first
        // and "the whole block" the rewrite names is no longer one paragraph.
        case "mergeOwn":
          return right.joinsNow;
        default:
          return true;
      }
    case "mergeOwn":
      switch (right.kind) {
        case "text":
        case "annotation":
        case "paragraphProperties":
        // A chain: this block's merge and the one merging into it.
        case "mergeNext":
          return false;
        default:
          return true;
      }
    case "mergeNext":
      // Applied directly, the join keeps the first paragraph's properties
      // and discards the joined one's: setting them edits a paragraph that
      // goes away. Tracked, the joined paragraph's mark is the one that
      // survives acceptance, so its properties are the merged paragraph's.
      // Its text, its own break and its own merge (a chain) survive either way.
      return right.kind === "paragraphProperties" && left.joinsNow;
    case "paragraphProperties":
      // Two property edits of one paragraph run from the end backwards, so the
      // earlier one would land last and override the later one; tracked, the
      // first to land leaves a pending property change the other cannot stack
      // on. The later one is refused, in either mode.
      return right.kind === "split" || right.kind === "paragraphProperties";
    case "split":
    case "text":
    case "annotation": {
      if (right.kind === "paragraphProperties") {
        return left.kind === "split";
      }
      if (right.kind !== "split" && right.kind !== "text" && right.kind !== "annotation") {
        // Every other kind was turned to the left above.
        return true;
      }
      // One paragraph break per block and batch: two splits would each set
      // the properties of a half the other one moves.
      if (left.kind === "split" && right.kind === "split") {
        return true;
      }
      // Marks move no position, so two of them never invalidate each other.
      if (left.kind === "annotation" && right.kind === "annotation") {
        return false;
      }
      // Annotations apply first, to the text as it was read (see the batch's
      // execution order), so one around an edit is an annotation of the text
      // the edit then rewrites, and the edit carries it as it carries any mark
      // around it. An annotation cut by an edit, or inside one, has no such
      // reading: part of what it names is gone.
      if (left.kind === "annotation" && contains(left.claim, right.claim)) {
        return false;
      }
      if (right.kind === "annotation" && contains(right.claim, left.claim)) {
        return false;
      }
      return inlineRangesConflict(left.claim, right.claim);
    }
  }
};

const insideAny = (ranges: readonly PositionRange[], position: number): boolean =>
  ranges.some(({ from, to }) => from <= position && position < to);

const strictlyInsideAny = (ranges: readonly PositionRange[], position: number): boolean =>
  ranges.some(({ from, to }) => from < position && position < to);

type TableRemovalClaim = Extract<BatchClaim, { type: "tableRemoval" }>;

/**
 * The claims of the operations a batch has accepted so far, indexed so a
 * batch of thousands of operations checks each one against its neighbours
 * rather than against every other operation.
 */
export class BatchClaims {
  private readonly blocks = new Map<number, BlockEntry[]>();
  private readonly insertions = new Map<number, string>();
  private readonly joins = new Map<number, string>();
  private readonly removals: { operationId: string; claim: TableRemovalClaim }[] = [];
  private readonly tables = new Map<number, string>();
  private readonly rowInsertions = new Map<number, string>();
  private readonly columnInsertions = new Map<number, string>();
  /** Accepted deletions that keep their paragraph, by where the paragraph ends. */
  private readonly keptParagraphs = new Map<number, { block: number; operationId: string }>();

  /** `claim`'s roles, given the insertions accepted so far. */
  private rolesOf(claim: BatchClaim): { block: number; role: BlockRole }[] {
    return blockRolesOf(claim).map(({ block, role }) =>
      role.kind === "deleteBlock" &&
      role.keepsParagraph &&
      claim.type === "deleteBlock" &&
      this.insertions.has(claim.end)
        ? { block, role: followedByInsertion(role) }
        : { block, role },
    );
  }

  /**
   * The operation an insertion at `at` conflicts with through a deletion it
   * would stop keeping its paragraph: one the deleted paragraph then no
   * longer allows.
   */
  private conflictOfFollowing(at: number): string | null {
    const kept = this.keptParagraphs.get(at);
    if (kept === undefined) {
      return null;
    }
    const entries = this.blocks.get(kept.block) ?? [];
    const deletion = entries.find(({ operationId }) => operationId === kept.operationId);
    if (deletion?.role.kind !== "deleteBlock") {
      return null;
    }
    const followed = followedByInsertion(deletion.role);
    return (
      entries.find(
        ({ operationId, role }) =>
          operationId !== kept.operationId && rolesConflict(followed, role),
      )?.operationId ?? null
    );
  }

  /**
   * The id of an accepted operation `claim` conflicts with, or `null` when it
   * conflicts with none.
   */
  conflictOf(claim: BatchClaim): string | null {
    for (const { block, role } of this.rolesOf(claim)) {
      for (const entry of this.blocks.get(block) ?? []) {
        if (rolesConflict(entry.role, role)) {
          return entry.operationId;
        }
      }
      // Deleting a block that a removed row, column or table takes with it
      // anyway removes nothing twice: the two deletions compose.
      if (role.kind === "deleteBlock") {
        continue;
      }
      for (const removal of this.removals) {
        if (insideAny(removal.claim.ranges, block)) {
          return removal.operationId;
        }
      }
    }
    switch (claim.type) {
      case "merge":
        return this.insertions.get(claim.next) ?? null;
      case "insertion": {
        const join = this.joins.get(claim.at);
        if (join !== undefined) {
          return join;
        }
        const following = this.conflictOfFollowing(claim.at);
        if (following !== null) {
          return following;
        }
        const removal = this.removals.find(({ claim: removed }) =>
          strictlyInsideAny(removed.ranges, claim.at),
        );
        return removal?.operationId ?? null;
      }
      case "tableStructure":
        return (
          this.removals.find(
            ({ claim: removed }) =>
              insideAny(removed.ranges, claim.table) ||
              (removed.table === claim.table &&
                removed.axis !== claim.axis &&
                (removed.axis === "row" || removed.axis === "column")),
          )?.operationId ??
          this.removals.find(
            ({ claim: removed }) => removed.wholeTable && removed.table === claim.table,
          )?.operationId ??
          null
        );
      case "tableRemoval": {
        if (claim.axis === "row") {
          const columnInsertion = this.columnInsertions.get(claim.table);
          if (columnInsertion !== undefined) return columnInsertion;
        }
        if (claim.axis === "column") {
          const rowInsertion = this.rowInsertions.get(claim.table);
          if (rowInsertion !== undefined) return rowInsertion;
        }
        if (claim.wholeTable) {
          const structural = this.tables.get(claim.table);
          if (structural !== undefined) {
            return structural;
          }
        }
        for (const removal of this.removals) {
          if (
            insideAny(removal.claim.ranges, claim.table) ||
            insideAny(claim.ranges, removal.claim.table) ||
            (removal.claim.table === claim.table &&
              (removal.claim.wholeTable || claim.wholeTable || removal.claim.axis !== claim.axis))
          ) {
            return removal.operationId;
          }
        }
        for (const [table, operationId] of this.tables) {
          if (strictlyInsideAny(claim.ranges, table)) {
            return operationId;
          }
        }
        for (const [block, entries] of this.blocks) {
          const editor = entries.find(({ role }) => role.kind !== "deleteBlock");
          if (editor !== undefined && insideAny(claim.ranges, block)) {
            return editor.operationId;
          }
        }
        for (const [at, operationId] of this.insertions) {
          if (strictlyInsideAny(claim.ranges, at)) {
            return operationId;
          }
        }
        return null;
      }
      default:
        return null;
    }
  }

  add(operationId: string, claim: BatchClaim): void {
    for (const { block, role } of this.rolesOf(claim)) {
      const entries = this.blocks.get(block);
      if (entries) {
        entries.push({ operationId, role });
      } else {
        this.blocks.set(block, [{ operationId, role }]);
      }
      if (claim.type === "deleteBlock" && role.kind === "deleteBlock" && role.keepsParagraph) {
        this.keptParagraphs.set(claim.end, { block, operationId });
      }
    }
    switch (claim.type) {
      case "merge":
        this.joins.set(claim.next, operationId);
        break;
      case "insertion": {
        if (!this.insertions.has(claim.at)) {
          this.insertions.set(claim.at, operationId);
        }
        const kept = this.keptParagraphs.get(claim.at);
        if (kept !== undefined) {
          this.keptParagraphs.delete(claim.at);
          for (const entry of this.blocks.get(kept.block) ?? []) {
            if (entry.operationId === kept.operationId && entry.role.kind === "deleteBlock") {
              entry.role = followedByInsertion(entry.role);
            }
          }
        }
        break;
      }
      case "tableStructure":
        if (!this.tables.has(claim.table)) {
          this.tables.set(claim.table, operationId);
        }
        if (claim.axis === "row" && !this.rowInsertions.has(claim.table)) {
          this.rowInsertions.set(claim.table, operationId);
        }
        if (claim.axis === "column" && !this.columnInsertions.has(claim.table)) {
          this.columnInsertions.set(claim.table, operationId);
        }
        break;
      case "tableRemoval":
        this.removals.push({ operationId, claim });
        if (!this.tables.has(claim.table)) {
          this.tables.set(claim.table, operationId);
        }
        break;
      default:
        break;
    }
  }
}
