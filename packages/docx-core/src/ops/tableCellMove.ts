import {
  type DocxPackage,
  MAX_REVISION_ID,
  type Paragraph,
  type ParagraphContent,
  type Run,
  type Table,
} from "../model/document";
import type { TableGeometryMove, TableGridChange } from "../model/formatting";
import { idKey, packageMoveRangeCensus } from "./ids";
import { stampInfo } from "./review";
import type { RevisionStamp } from "./types";

export type TableMoveSource = { paragraph: Paragraph; id: string; runs: Run[] };

/** The run move can take the paragraph whole without duplicating an anchor. */
export const tableMoveSource = (
  paragraph: Paragraph,
): TableMoveSource | undefined => {
  if (
    paragraph.paraId === undefined ||
    paragraph.content.length === 0 ||
    paragraph.pPrMark !== undefined ||
    paragraph.propertyChanges !== undefined ||
    paragraph.sectionProperties !== undefined ||
    paragraph.reviewCarrier !== undefined
  )
    return undefined;
  const runs: Run[] = [];
  for (const item of paragraph.content) {
    if (
      item.type !== "run" ||
      item.propertyChanges !== undefined ||
      !item.content.every((leaf) => leaf.type === "text")
    )
      return undefined;
    runs.push(item);
  }
  return { paragraph, id: paragraph.paraId, runs };
};

type MarkTableMoveOptions = {
  source: TableMoveSource;
  destination: string;
  from: RevisionStamp;
  to: RevisionStamp;
  mark: RevisionStamp;
  ranges: { fromRange: number; toRange: number; name: string };
};
export const moveRangeAllocator = (pkg: DocxPackage) => {
  const { ids, names } = packageMoveRangeCensus(pkg);
  let nextId = 1;
  let nextName = 1;
  const takeId = () => {
    while (ids.has(nextId) && nextId < MAX_REVISION_ID) nextId++;
    return nextId < MAX_REVISION_ID ? nextId++ : undefined;
  };
  return () => {
    const fromRange = takeId();
    const toRange = takeId();
    if (fromRange === undefined || toRange === undefined) return undefined;
    let name = `FolioMove${nextName++}`;
    while (names.has(name)) name = `FolioMove${nextName++}`;
    return { fromRange, toRange, name };
  };
};

export const markTableMove = ({
  source,
  destination,
  from,
  to,
  mark,
  ranges,
}: MarkTableMoveOptions) => {
  const { fromRange, toRange, name } = ranges;
  const record = {
    source: source.id,
    destination,
    fromRevision: from.id,
    toRevision: to.id,
    fromRange,
    toRange,
    destinationMark: mark.id,
  } satisfies TableGeometryMove;
  return {
    record,
    source: {
      ...source.paragraph,
      content: [
        {
          type: "moveFromRangeStart",
          id: fromRange,
          name,
          author: from.author,
          date: from.date,
        },
        { type: "moveFrom", info: stampInfo(from), content: source.runs },
        { type: "moveFromRangeEnd", id: fromRange },
      ],
    } satisfies Paragraph,
    destination: {
      ...source.paragraph,
      paraId: destination,
      pPrMark: { kind: "moveTo", info: stampInfo(mark) },
      content: [
        {
          type: "moveToRangeStart",
          id: toRange,
          name,
          author: to.author,
          date: to.date,
        },
        { type: "moveTo", info: stampInfo(to), content: source.runs },
        { type: "moveToRangeEnd", id: toRange },
      ],
    } satisfies Paragraph,
  };
};

export const tableMoveRevisionIds = (
  moves: readonly TableGeometryMove[],
): number[] =>
  moves.flatMap(({ fromRevision, toRevision, destinationMark }) => [
    fromRevision,
    toRevision,
    destinationMark,
  ]);

type MergeGeometry = Extract<
  NonNullable<TableGridChange["geometryChange"]>,
  { type: "mergeCells" }
>;

/** Check the move's saved anchors against the native records in this merge. */
export const tableMoveOwnershipMatches = (
  table: Table,
  geometry: MergeGeometry,
): boolean => {
  if (geometry.moves === undefined && geometry.emptyOwner === undefined)
    return true;
  const members = new Set(geometry.revisionIds);
  const sources = new Map<string, Paragraph>();
  const destinations = new Map<string, Paragraph>();
  let hasOwner = false;
  let emptyOwner: string | undefined;
  for (const row of table.rows) {
    const removedRow =
      row.structuralChange?.type === "tableRowDeletion" &&
      members.has(row.structuralChange.info.id);
    for (const cell of row.cells) {
      const removedCell =
        cell.structuralChange?.type === "tableCellDeletion" &&
        members.has(cell.structuralChange.info.id);
      const owned =
        cell.propertyChanges?.some(({ info }) => members.has(info.id)) ?? false;
      if (!owned && !removedRow && !removedCell) continue;
      if (owned) {
        if (hasOwner || removedRow || removedCell) return false;
        hasOwner = true;
      }
      const first = cell.content.at(0);
      if (owned && first?.type === "paragraph") emptyOwner = first.paraId;
      const paragraphs = owned ? destinations : sources;
      for (const paragraph of cell.content) {
        if (paragraph.type === "paragraph" && paragraph.paraId !== undefined)
          paragraphs.set(idKey(paragraph.paraId), paragraph);
      }
    }
  }
  if (
    geometry.emptyOwner !== undefined &&
    (emptyOwner === undefined ||
      idKey(geometry.emptyOwner) !== idKey(emptyOwner) ||
      (geometry.moves?.length ?? 0) === 0)
  )
    return false;
  const seen = new Set<string>();
  for (const move of geometry.moves ?? []) {
    const sourceId = idKey(move.source);
    const destinationId = idKey(move.destination);
    if (
      sourceId === destinationId ||
      seen.has(sourceId) ||
      seen.has(destinationId)
    )
      return false;
    seen.add(sourceId);
    seen.add(destinationId);
    const source = sources.get(sourceId);
    const destination = destinations.get(destinationId);
    if (
      source?.content.length !== 3 ||
      destination?.content.length !== 3 ||
      destination.pPrMark?.kind !== "moveTo" ||
      destination.pPrMark.info.id !== move.destinationMark
    )
      return false;
    const [fromStart, from, fromEnd] = source.content;
    const [toStart, to, toEnd] = destination.content;
    if (
      fromStart?.type !== "moveFromRangeStart" ||
      fromStart.id !== move.fromRange ||
      from?.type !== "moveFrom" ||
      from.info.id !== move.fromRevision ||
      fromEnd?.type !== "moveFromRangeEnd" ||
      fromEnd.id !== move.fromRange ||
      toStart?.type !== "moveToRangeStart" ||
      toStart.id !== move.toRange ||
      to?.type !== "moveTo" ||
      to.info.id !== move.toRevision ||
      toEnd?.type !== "moveToRangeEnd" ||
      toEnd.id !== move.toRange ||
      fromStart.name !== toStart.name
    )
      return false;
  }
  return true;
};

/** Settle paragraph identity at the same time as the cell topology, keeping
 * the source id at the destination on acceptance. */
export const resolveTableMoves = (table: Table, reject: boolean): Table => {
  const geometry = table.formatting?.gridChange?.geometryChange;
  if (geometry?.type !== "mergeCells" || geometry.moves === undefined)
    return table;
  const destinations = new Map(
    geometry.moves.map((move) => [idKey(move.destination), move]),
  );
  const sources = new Map(
    geometry.moves.map((move) => [idKey(move.source), move]),
  );
  const emptyOwner =
    geometry.emptyOwner === undefined ? undefined : idKey(geometry.emptyOwner);
  const rows = table.rows.map((row) =>
    Object.assign({}, row, {
      cells: row.cells.map((cell) =>
        Object.assign({}, cell, {
          content: cell.content.flatMap((block) => {
            if (block.type !== "paragraph" || block.paraId === undefined)
              return [block];
            const key = idKey(block.paraId);
            const destination = destinations.get(key);
            if (
              (reject && destination !== undefined) ||
              (!reject &&
                key === emptyOwner &&
                block.content.length === 0 &&
                block.pPrMark === undefined &&
                block.propertyChanges === undefined &&
                block.sectionProperties === undefined)
            )
              return [];
            const move = reject ? sources.get(key) : destination;
            if (move === undefined) return [block];
            const revision = reject ? move.fromRevision : move.toRevision;
            const range = reject ? move.fromRange : move.toRange;
            const content = block.content.flatMap(
              (item): ParagraphContent[] => {
                switch (item.type) {
                  case "moveFromRangeStart":
                  case "moveFromRangeEnd":
                  case "moveToRangeStart":
                  case "moveToRangeEnd":
                    return item.id === range ? [] : [item];
                  case "moveFrom":
                  case "moveTo":
                    return item.info.id === revision ? item.content : [item];
                  default:
                    return [item];
                }
              },
            );
            const paragraph = { ...block, paraId: move.source, content };
            if (paragraph.pPrMark?.info.id === move.destinationMark)
              delete paragraph.pPrMark;
            return [paragraph];
          }),
        }),
      ),
    }),
  );
  return { ...table, rows };
};
