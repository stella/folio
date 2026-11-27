import type { BlockContent, Table } from "./content";
import type { TableGeometryNestedMove } from "./formatting";

const tableBlocks = (table: Table) =>
  table.rows.flatMap(({ cells }) => cells.flatMap(({ content }) => content));

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

const propertySlots = (table: Table): ReadonlyMap<number, string> => {
  const slots = new Map<number, string>();
  const record = (
    changes: readonly { type: string; info: { id: number } }[] | undefined,
    position: string,
  ) => {
    for (const [index, { type, info }] of (changes ?? []).entries())
      slots.set(info.id, `${position}:${type}:${index}`);
  };
  const owners = [table];
  visitBlocks(tableBlocks(table), (block) => {
    if (block.type === "table") owners.push(block);
  });
  for (const [index, owner] of owners.entries()) {
    record(owner.propertyChanges, `${index}:table`);
    for (const [rowIndex, row] of owner.rows.entries()) {
      const position = `${index}:row:${rowIndex}`;
      record(row.propertyChanges, position);
      record(row.tablePropertyExceptionChanges, position);
      for (const [cellIndex, cell] of row.cells.entries())
        record(cell.propertyChanges, `${position}:cell:${cellIndex}`);
    }
  }
  return slots;
};

/** Active property pairs must belong to the two recorded nested tables. */
export const activeNestedTablePropertyHandoffs = (
  owner: Table,
  move: TableGeometryNestedMove,
): readonly (readonly [number, number])[] => {
  if (typeof move !== "object" || move === null) return [];
  const { properties, paragraphs, rows } = move;
  if (
    !Array.isArray(properties) ||
    properties.length === 0 ||
    !Array.isArray(paragraphs) ||
    paragraphs.length === 0 ||
    !Array.isArray(rows)
  )
    return [];
  const tables: Table[] = [];
  visitBlocks(tableBlocks(owner), (block) => {
    if (block.type === "table") tables.push(block);
  });
  const inserted = new Set(rows);
  const target = tables.find(({ rows: candidateRows }) =>
    candidateRows.some(
      ({ structuralChange }) =>
        structuralChange?.type === "tableRowInsertion" &&
        inserted.has(structuralChange.info.id),
    ),
  );
  if (target === undefined) return [];
  const destinations = paragraphIds(target);
  const pairs: [string, string][] = [];
  for (const pair of paragraphs) {
    if (!Array.isArray(pair) || pair.length !== 2) return [];
    const [source, destination] = pair;
    if (typeof source !== "string" || typeof destination !== "string")
      return [];
    const from = source.toUpperCase();
    const to = destination.toUpperCase();
    if (destinations.has(from) || !destinations.has(to)) return [];
    pairs.push([from, to]);
  }
  const source = tables.find((candidate) => {
    if (candidate === target) return false;
    const origins = paragraphIds(candidate);
    return pairs.every(([from, to]) => origins.has(from) && !origins.has(to));
  });
  if (source === undefined) return [];
  const originals = propertySlots(source);
  const copies = propertySlots(target);
  return properties.filter((pair) => {
    if (!Array.isArray(pair) || pair.length !== 2) return false;
    const [old, temporary] = pair;
    const slot = originals.get(old);
    return slot !== undefined && slot === copies.get(temporary);
  });
};
