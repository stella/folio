import type { BlockContent, Table } from "./content";
import type { TableGeometryNestedMove } from "./formatting";

const tableBlocks = (table: Table) =>
  table.rows.flatMap(({ cells }) => cells.flatMap(({ content }) => content));

const directTableParagraphIds = (table: Table): Set<string> => {
  const ids = new Set<string>();
  const visit = (blocks: readonly BlockContent[]): void => {
    for (const block of blocks) {
      if (block.type === "paragraph" && block.paraId !== undefined)
        ids.add(block.paraId.toUpperCase());
      else if (block.type === "blockCustomXml" || block.type === "blockSdt")
        visit(block.content);
    }
  };
  visit(tableBlocks(table));
  return ids;
};

const visitBlocks = (
  blocks: readonly BlockContent[],
  visit: (block: BlockContent) => void,
): void => {
  for (const block of blocks) {
    if (block.type === "table") visitBlocks(tableBlocks(block), visit);
    else if (block.type === "blockCustomXml" || block.type === "blockSdt")
      visitBlocks(block.content, visit);
    visit(block);
  }
};

const paragraphIds = (table: Table): Set<string> => {
  const ids = new Set<string>();
  visitBlocks(tableBlocks(table), (block) => {
    if (block.type === "paragraph" && block.paraId !== undefined)
      ids.add(block.paraId.toUpperCase());
  });
  return ids;
};

const nestedTables = (owner: Table): Table[] => {
  const tables: Table[] = [];
  visitBlocks(tableBlocks(owner), (block) => {
    if (block.type === "table") tables.push(block);
  });
  return tables;
};

/** An authored-only row copy keeps its destination anchors after review. */
export const nestedTableCopyByAnchors = (
  owner: Table,
  paragraphs: TableGeometryNestedMove["paragraphs"],
): Table | undefined => {
  const live = paragraphIds(owner);
  const destinations = new Set(
    paragraphs.map(([, to]) => to.toUpperCase()).filter((id) => live.has(id)),
  );
  if (destinations.size === 0) return undefined;
  return nestedTables(owner).find((table) => {
    const direct = directTableParagraphIds(table);
    if (![...direct].some((id) => destinations.has(id))) return false;
    const inside = paragraphIds(table);
    return (
      [...destinations].every((id) => inside.has(id)) &&
      paragraphs.every(([from]) => !inside.has(from.toUpperCase()))
    );
  });
};

/** Property anchors only belong to tables containing no cells. */
export const omittedTableAnchorId = (table: Table): number | undefined => {
  const id = table.propertyChanges?.at(0)?.info.id;
  return id !== undefined &&
    table.rows.length > 0 &&
    table.rows.every(({ cells }) => cells.length === 0)
    ? id
    : undefined;
};

const revisionSlots = (table: Table): ReadonlyMap<number, string> => {
  const slots = new Map<number, string>();
  const record = (
    changes: readonly { type: string; info: { id: number } }[] | undefined,
    position: string,
  ) => {
    for (const [index, { type, info }] of (changes ?? []).entries())
      slots.set(info.id, `${position}:${type}:${index}`);
  };
  const owners = [table];
  let paragraphIndex = 0;
  visitBlocks(tableBlocks(table), (block) => {
    if (block.type === "table") owners.push(block);
    if (block.type !== "paragraph") return;
    const position = `paragraph:${paragraphIndex++}`;
    const mark = block.pPrMark;
    if (mark !== undefined) slots.set(mark.info.id, `${position}:${mark.kind}`);
    const content = block.content.at(1);
    if (content?.type === "moveFrom" || content?.type === "moveTo")
      slots.set(content.info.id, `${position}:${content.type}`);
  });
  for (const [index, owner] of owners.entries()) {
    const grid = owner.formatting?.gridChange;
    if (grid !== undefined)
      slots.set(
        grid.id,
        `${index}:grid:${grid.geometryChange?.type ?? "plain"}`,
      );
    record(owner.propertyChanges, `${index}:table`);
    for (const [rowIndex, row] of owner.rows.entries()) {
      const position = `${index}:row:${rowIndex}`;
      if (row.structuralChange !== undefined)
        slots.set(
          row.structuralChange.info.id,
          `${position}:${row.structuralChange.type}`,
        );
      record(row.propertyChanges, position);
      record(row.tablePropertyExceptionChanges, position);
      for (const [cellIndex, cell] of row.cells.entries()) {
        const cellPosition = `${position}:cell:${cellIndex}`;
        record(cell.propertyChanges, cellPosition);
        for (const [snapshotIndex, { previousStructuralChange }] of (
          cell.propertyChanges ?? []
        ).entries()) {
          if (previousStructuralChange !== undefined)
            slots.set(
              previousStructuralChange.info.id,
              `${cellPosition}:previous:${snapshotIndex}:${previousStructuralChange.type}`,
            );
        }
        if (cell.structuralChange !== undefined)
          slots.set(
            cell.structuralChange.info.id,
            `${cellPosition}:${cell.structuralChange.type}`,
          );
      }
    }
  }
  return slots;
};

/** Active revision pairs must belong to the two recorded nested tables. */
export const activeNestedTableRevisionHandoffs = (
  owner: Table,
  move: TableGeometryNestedMove,
): readonly (readonly [number, number])[] => {
  if (typeof move !== "object" || move === null) return [];
  const { revisions, paragraphs, rows } = move;
  if (
    !Array.isArray(revisions) ||
    revisions.length === 0 ||
    !Array.isArray(paragraphs) ||
    !Array.isArray(rows)
  )
    return [];
  const tables = nestedTables(owner);
  const pairs: [string, string][] = [];
  for (const pair of paragraphs) {
    if (!Array.isArray(pair) || pair.length !== 2) return [];
    const [source, destination] = pair;
    if (typeof source !== "string" || typeof destination !== "string")
      return [];
    const from = source.toUpperCase();
    const to = destination.toUpperCase();
    pairs.push([from, to]);
  }
  const inserted = new Set(rows);
  const omittedTable = move.omittedTable;
  let target: Table | undefined;
  if (omittedTable !== undefined) {
    target = tables.find(
      (candidate) => omittedTableAnchorId(candidate) === omittedTable[1],
    );
  } else if (inserted.size === 0) {
    target = nestedTableCopyByAnchors(owner, pairs);
  } else {
    target = tables.find(({ rows: candidateRows }) =>
      candidateRows.some(
        ({ structuralChange }) =>
          structuralChange?.type === "tableRowInsertion" &&
          inserted.has(structuralChange.info.id),
      ),
    );
  }
  if (target === undefined) return [];
  const copies = revisionSlots(target);
  const pairedIn = (candidate: Table) => {
    const originals = revisionSlots(candidate);
    return revisions.filter((pair) => {
      if (!Array.isArray(pair) || pair.length !== 2) return false;
      const [old, temporary] = pair;
      const slot = originals.get(old);
      return slot !== undefined && slot === copies.get(temporary);
    });
  };
  if (omittedTable !== undefined) {
    const source = tables.find(
      (candidate) => omittedTableAnchorId(candidate) === omittedTable[0],
    );
    return source === undefined || source === target ? [] : pairedIn(source);
  }
  // Fully omitted rows have no paragraphs to pair. A native insertion
  // identifies the target. Matching physical slots identify the source
  // while both versions still carry the earlier record.
  if (pairs.length === 0) {
    for (const candidate of tables) {
      if (candidate === target) continue;
      const active = pairedIn(candidate);
      if (active.length > 0) return active;
    }
    return [];
  }
  const destinations = paragraphIds(target);
  if (pairs.some(([from]) => destinations.has(from))) return [];
  // A decision on an earlier cell record may have removed the same
  // paragraphs from both copies before this move is decided.
  if (!pairs.some(([, to]) => destinations.has(to))) return [];
  const source = tables.find((candidate) => {
    if (candidate === target) return false;
    const origins = paragraphIds(candidate);
    return pairs.every(
      ([from, to]) =>
        !origins.has(to) && origins.has(from) === destinations.has(to),
    );
  });
  return source === undefined ? [] : pairedIn(source);
};
