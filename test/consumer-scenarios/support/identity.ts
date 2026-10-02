/**
 * Paragraph IDs without authored package IDs are allocated from position and
 * may be recomputed when tracked text is resolved. Compare those rows by
 * their position while keeping authored IDs and all content strict.
 */

import type { openReviewer } from "./documents.ts";

type Reviewer = Awaited<ReturnType<typeof openReviewer>>;
type ContentRow = ReturnType<Reviewer["getContent"]>[number];
type Snapshot = ReturnType<Reviewer["snapshot"]>;

type IdentifiedRow = Pick<ContentRow, "id" | "idStability">;
type IdentityPairOptions<Left extends IdentifiedRow, Right extends IdentifiedRow> = {
  leftRows: readonly Left[];
  rightRows: readonly Right[];
  /** Stable authored ids from the live reviewer, including rows outside a saved baseline. */
  stableIds?: ReadonlySet<string>;
};

const tokenAt = (index: number): string => `\u0000position:${String(index)}`;
const countIds = (rows: readonly IdentifiedRow[]): Map<string, number> => {
  const counts = new Map<string, number>();
  for (const { id } of rows) counts.set(id, (counts.get(id) ?? 0) + 1);
  return counts;
};
const withoutIdStability = <Row extends IdentifiedRow>(row: Row) => {
  const projected = { ...row };
  delete projected.idStability;
  return projected;
};

const identityAliases = <Left extends IdentifiedRow, Right extends IdentifiedRow>({
  leftRows,
  rightRows,
  stableIds = new Set(),
}: IdentityPairOptions<Left, Right>) => {
  const stable = new Set(stableIds);
  for (const row of leftRows) {
    if (row.idStability !== "positional") stable.add(row.id);
  }
  const leftCounts = countIds(leftRows);
  const rightCounts = countIds(rightRows);
  const leftAliases = new Map<string, string>();
  const rightAliases = new Map<string, string>();
  const leftPositional = new Set<number>();
  const rightPositional = new Set<number>();

  for (const [index, left] of leftRows.entries()) {
    if (left.idStability !== "positional") continue;
    const right = rightRows[index];
    if (!right) continue;
    if (leftCounts.get(left.id) === 1 && !stable.has(left.id)) {
      leftAliases.set(left.id, tokenAt(index));
      leftPositional.add(index);
    }
    if (rightCounts.get(right.id) === 1 && !stable.has(right.id)) {
      rightAliases.set(right.id, tokenAt(index));
      rightPositional.add(index);
    }
  }
  return { leftAliases, rightAliases, leftPositional, rightPositional };
};

/** Compare content rows while treating only paired positional identity as ordinal. */
export const projectContentPair = <Left extends IdentifiedRow, Right extends IdentifiedRow>({
  leftRows,
  rightRows,
  stableIds,
}: IdentityPairOptions<Left, Right>) => {
  const identity = identityAliases({ leftRows, rightRows, stableIds });
  const left = leftRows.map((row, index) => {
    if (!identity.leftPositional.has(index)) return row;
    return { ...withoutIdStability(row), id: identity.leftAliases.get(row.id) ?? row.id };
  });
  const right = rightRows.map((row, index) => {
    if (!identity.rightPositional.has(index)) return row;
    return { ...withoutIdStability(row), id: identity.rightAliases.get(row.id) ?? row.id };
  });
  return { left, right, ...identity };
};

/** The bridge snapshot exposes block ids in blocks, anchor keys and anchor.id. */
export const projectSnapshotIdentities = (
  snapshot: Snapshot,
  aliases: ReadonlyMap<string, string>,
): Snapshot => {
  const blocks = snapshot.blocks.map((block) => {
    const positionalId = aliases.get(block.id);
    if (positionalId === undefined) return block;
    const projected = { ...block, id: positionalId };
    Reflect.deleteProperty(projected, "idStability");
    return projected;
  });
  const anchors = { ...snapshot.anchors };
  for (const [key, anchor] of Object.entries(snapshot.anchors)) {
    const mappedKey = aliases.get(key) ?? key;
    const mappedId = aliases.get(anchor.id) ?? anchor.id;
    Reflect.deleteProperty(anchors, key);
    anchors[mappedKey] = { ...anchor, id: mappedId };
  }
  return { ...snapshot, blocks, anchors };
};
