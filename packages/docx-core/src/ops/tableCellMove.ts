import { isNoteReferenceId } from "../model/content";
import {
  type BlockContent,
  type DocxPackage,
  type InlineSdt,
  MAX_REVISION_ID,
  type Paragraph,
  type ParagraphContent,
  type Table,
} from "../model/document";
import type {
  TableGeometryMove,
  TableGeometryNestedMove,
  TableGridChange,
} from "../model/formatting";
import { readXml } from "../xml/streamingXmlParser";
import { OOXML_NAMESPACE_SCOPE, OOXML_NS } from "../xml/xmlNamespaceContext";
import {
  identityKeysIn,
  idKey,
  packageMoveRangeCensus,
  paragraphIdsIn,
} from "./ids";
import {
  asParagraphContent,
  childNodes,
  type InlineNode,
  rebuildNode,
} from "./leaves";
import { stampInfo } from "./review";
import { reachesPastRemovedCell } from "./tableEditDependencies";
import type { RevisionStamp } from "./types";

export type TableMoveSource = {
  paragraph: Paragraph;
  id: string;
  controls: number[];
};
export type NestedTableMoveSource = { table: Table; paragraphs: string[] };

const movableAtom = (leaf: RunContent): boolean =>
  leaf.type === "text" ||
  leaf.type === "tab" ||
  (leaf.type === "break" &&
    (leaf.breakType === undefined || leaf.breakType === "textWrapping")) ||
  leaf.type === "symbol" ||
  leaf.type === "softHyphen" ||
  leaf.type === "noBreakHyphen" ||
  ((leaf.type === "footnoteRef" || leaf.type === "endnoteRef") &&
    Number.isInteger(leaf.id) &&
    isNoteReferenceId(leaf.id));

const movableRun = (run: Run): boolean =>
  run.propertyChanges === undefined && run.content.every(movableAtom);

const emptyWordElement = (xml: string, localName: string): boolean => {
  const parsed = readXml(xml, {
    conformance: "ooxml",
    inheritedNamespaceScope: OOXML_NAMESPACE_SCOPE,
  });
  if (parsed.status !== "parsed" || parsed.value.elements?.length !== 1)
    return false;
  const element = parsed.value.elements.at(0);
  return (
    element?.type === "element" &&
    (element.namespaceUri === OOXML_NS.w ||
      element.namespaceUri ===
        "http://purl.oclc.org/ooxml/wordprocessingml/main") &&
    element.name?.slice(element.name.indexOf(":") + 1) === localName &&
    (element.elements?.length ?? 0) === 0 &&
    Object.keys(element.attributes ?? {}).every(
      (name) => name === "xmlns" || name.startsWith("xmlns:"),
    )
  );
};

/** Captured properties can carry binding or range data that a copy cannot take. */
const movableControlId = (
  properties: InlineSdt["properties"],
): number | undefined => {
  const { id, preserved, rawEndPropertiesXml } = properties;
  if (
    properties.sdtType !== "richText" ||
    id === undefined ||
    !Number.isInteger(id) ||
    id < -MAX_REVISION_ID - 1 ||
    id > MAX_REVISION_ID ||
    properties.rawSdtChildrenBeforeContent !== undefined ||
    properties.rawSdtChildrenAfterContent !== undefined ||
    properties.endProperties?.runProperties !== undefined ||
    (rawEndPropertiesXml !== undefined &&
      !emptyWordElement(rawEndPropertiesXml, "sdtEndPr"))
  )
    return undefined;
  const propertiesXml = preserved?.children ?? [];
  return propertiesXml.length === 0 ||
    (propertiesXml.length === 1 &&
      propertiesXml.every(({ xml }) => emptyWordElement(xml, "richText")))
    ? id
    : undefined;
};

type MoveContentContext = { withinLink: boolean; controls: number[] };

// Both sides of the move carry the inline content. Admit only records whose
// identity and ownership stay local to the paragraph.
const movableInlineContent = (
  content: readonly ParagraphContent[],
  context: MoveContentContext,
): boolean =>
  content.length > 0 &&
  content.every((item) => {
    switch (item.type) {
      case "run":
        return movableRun(item);
      case "hyperlink":
        // A bidi wrapper can contain a link, but cannot hide a nested one.
        return (
          !context.withinLink &&
          movableInlineContent(item.children, { ...context, withinLink: true })
        );
      case "inlineWrapper":
        return (
          item.kind === "bidi" && movableInlineContent(item.content, context)
        );
      case "inlineSdt": {
        const id = movableControlId(item.properties);
        if (id === undefined) return false;
        context.controls.push(id);
        return movableInlineContent(item.content, context);
      }
      default:
        return false;
    }
  });

/** The run move can take the paragraph whole without duplicating an anchor. */
export const tableMoveSource = (
  paragraph: Paragraph,
): TableMoveSource | undefined => {
  const controls: number[] = [];
  if (
    paragraph.paraId === undefined ||
    paragraph.pPrMark !== undefined ||
    paragraph.propertyChanges !== undefined ||
    paragraph.sectionProperties !== undefined ||
    paragraph.reviewCarrier !== undefined ||
    (paragraph.content.length !== 0 &&
      !movableInlineContent(paragraph.content, { withinLink: false, controls }))
  )
    return undefined;
  return { paragraph, id: paragraph.paraId, controls };
};

/** A structural copy may carry only records whose package identities stay unique. */
export const nestedTableMoveSource = (
  table: Table,
): NestedTableMoveSource | undefined => {
  if (identityKeysIn(table).length !== 0 || reachesPastRemovedCell(table))
    return undefined;
  const paragraphs: string[] = [];
  const collect = (nested: Table): boolean =>
    nested.bookmarks === undefined &&
    nested.carrierStack === undefined &&
    nested.rows.length > 0 &&
    nested.rows.every(
      (row) =>
        row.bookmarks === undefined &&
        row.contentControls === undefined &&
        row.carrierStack === undefined &&
        row.cells.length > 0 &&
        row.cells.every(
          (cell) =>
            cell.contentControls === undefined &&
            cell.carrierStack === undefined &&
            cell.content.at(-1)?.type === "paragraph" &&
            cell.content.every((block) => {
              if (block.type === "table") return collect(block);
              if (block.type !== "paragraph") return false;
              const source = tableMoveSource(block);
              if (source === undefined) return false;
              paragraphs.push(source.id);
              return true;
            }),
        ),
    );
  return collect(table) ? { table, paragraphs } : undefined;
};

const withNestedParagraphIds = (
  table: Table,
  ids: ReadonlyMap<string, string>,
): Table => ({
  ...table,
  rows: table.rows.map((row) => ({
    ...row,
    cells: row.cells.map((cell) => ({
      ...cell,
      content: cell.content.map((block) => {
        if (block.type === "table") return withNestedParagraphIds(block, ids);
        if (block.type !== "paragraph" || block.paraId === undefined)
          return block;
        const paraId = ids.get(idKey(block.paraId));
        return paraId === undefined ? block : { ...block, paraId };
      }),
    })),
  })),
});

type MarkNestedTableMoveOptions = {
  source: NestedTableMoveSource;
  destinations: readonly string[];
  rowStamps: readonly RevisionStamp[];
};
type MarkedNestedTableMove = {
  destination: Table;
  record: TableGeometryNestedMove;
};
export const markNestedTableMove = ({
  source,
  destinations,
  rowStamps,
}: MarkNestedTableMoveOptions): MarkedNestedTableMove | undefined => {
  if (
    destinations.length !== source.paragraphs.length ||
    rowStamps.length !== source.table.rows.length
  )
    return undefined;
  const paragraphs: [string, string][] = [];
  for (const [index, original] of source.paragraphs.entries()) {
    const destination = destinations[index];
    if (destination === undefined) return undefined;
    paragraphs.push([original, destination]);
  }
  const table = withNestedParagraphIds(
    source.table,
    new Map(
      paragraphs.map(([original, destination]) => [
        idKey(original),
        destination,
      ]),
    ),
  );
  const rows = table.rows.map((row, index) => {
    const stamp = rowStamps[index];
    return stamp === undefined
      ? row
      : {
          ...row,
          structuralChange: {
            type: "tableRowInsertion" as const,
            info: stampInfo(stamp),
          },
        };
  });
  return {
    destination: { ...table, rows },
    record: {
      rows: rowStamps.map(({ id }) => id),
      paragraphs,
    } satisfies TableGeometryNestedMove,
  };
};

type TableMoveRanges = { fromRange: number; toRange: number; name: string };
type MarkTableMoveOptions = {
  source: TableMoveSource;
  destination: string;
  from: RevisionStamp;
  to: RevisionStamp;
  mark: RevisionStamp;
  ranges: TableMoveRanges;
  controls: readonly (readonly [number, number])[];
};

const withControlIds = (
  content: ParagraphContent[],
  ids: ReadonlyMap<number, number>,
): ParagraphContent[] => {
  if (ids.size === 0) return content;
  const rewrite = (node: InlineNode): InlineNode => {
    const children = childNodes(node);
    const rewritten = children?.map(rewrite);
    const next =
      rewritten?.some((child, index) => child !== children?.[index]) === true
        ? rebuildNode(node, rewritten)
        : node;
    if (next.type !== "inlineSdt" || next.properties.id === undefined)
      return next;
    const id = ids.get(next.properties.id);
    return id === undefined
      ? next
      : { ...next, properties: { ...next.properties, id } };
  };
  return asParagraphContent(content.map(rewrite));
};
export const moveRangeAllocator = (
  pkg: DocxPackage,
): (() => TableMoveRanges | undefined) => {
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

type MarkedTableMove = {
  record: TableGeometryMove;
  source: Paragraph;
  destination: Paragraph;
};
export const markTableMove = ({
  source,
  destination,
  from,
  to,
  mark,
  ranges,
  controls,
}: MarkTableMoveOptions): MarkedTableMove => {
  const { fromRange, toRange, name } = ranges;
  const record = {
    source: source.id,
    destination,
    fromRevision: from.id,
    toRevision: to.id,
    fromRange,
    toRange,
    destinationMark: mark.id,
    ...(controls.length === 0 ? {} : { controls }),
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
        {
          type: "moveFrom",
          info: stampInfo(from),
          content: source.paragraph.content,
        },
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
        {
          type: "moveTo",
          info: stampInfo(to),
          content: withControlIds(source.paragraph.content, new Map(controls)),
        },
        { type: "moveToRangeEnd", id: toRange },
      ],
    } satisfies Paragraph,
  };
};

export const markEmptyTableMove = ({
  source,
  destination,
  mark,
}: Pick<
  MarkTableMoveOptions,
  "source" | "destination" | "mark"
>): MarkedTableMove => ({
  record: {
    kind: "empty",
    source: source.id,
    destination,
    destinationMark: mark.id,
  } satisfies TableGeometryMove,
  source: source.paragraph,
  destination: {
    ...source.paragraph,
    paraId: destination,
    pPrMark: { kind: "ins", info: stampInfo(mark) },
  } satisfies Paragraph,
});

export const tableMoveRevisionIds = (
  moves: readonly TableGeometryMove[],
): number[] =>
  moves.flatMap((move) =>
    move.kind === "empty"
      ? [move.destinationMark]
      : [move.fromRevision, move.toRevision, move.destinationMark],
  );

type MergeGeometry = Extract<
  NonNullable<TableGridChange["geometryChange"]>,
  { type: "mergeCells" }
>;

/** Check the move's saved anchors against the native records in this merge. */
export const tableMoveOwnershipMatches = (
  table: Table,
  geometry: MergeGeometry,
): boolean => {
  if (
    geometry.moves === undefined &&
    geometry.nestedMoves === undefined &&
    geometry.emptyOwner === undefined
  )
    return true;
  const members = new Set(geometry.revisionIds);
  const sources = new Map<string, Paragraph>();
  const destinations = new Map<string, Paragraph>();
  const movedRows = new Map<number, Table>();
  const nestedSources = new Map<string, string[]>();
  const collectMovedTables = (
    blocks: readonly BlockContent[],
    side: "source" | "destination",
  ): void => {
    for (const block of blocks) {
      if (block.type === "table") {
        if (side === "source") {
          const ids = paragraphIdsIn(block).map(idKey);
          const first = ids.at(0);
          if (first !== undefined) nestedSources.set(first, ids);
          continue;
        }
        for (const { structuralChange } of block.rows) {
          if (structuralChange?.type === "tableRowInsertion")
            movedRows.set(structuralChange.info.id, block);
        }
      } else if (block.type === "blockCustomXml" || block.type === "blockSdt") {
        collectMovedTables(block.content, side);
      }
    }
  };
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
      if (geometry.nestedMoves !== undefined)
        collectMovedTables(cell.content, owned ? "destination" : "source");
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
      (geometry.moves?.length ?? 0) + (geometry.nestedMoves?.length ?? 0) === 0)
  )
    return false;
  const nestedTables = new Set<Table>();
  const claimedSources = new Set<string>();
  for (const move of geometry.nestedMoves ?? []) {
    const first = move.rows.at(0);
    const destination = first === undefined ? undefined : movedRows.get(first);
    if (
      destination === undefined ||
      nestedTables.has(destination) ||
      new Set(move.rows).size !== move.rows.length ||
      move.rows.some((id) => movedRows.get(id) !== destination)
    )
      return false;
    const firstSource = move.paragraphs.at(0)?.at(0);
    if (firstSource === undefined) return false;
    const sourceId = idKey(firstSource);
    const sourceIds = nestedSources.get(sourceId);
    const destinationIds = paragraphIdsIn(destination).map(idKey);
    if (
      claimedSources.has(sourceId) ||
      sourceIds?.length !== move.paragraphs.length ||
      destinationIds.length !== move.paragraphs.length ||
      move.paragraphs.some(
        ([from, to], index) =>
          sourceIds[index] !== idKey(from) ||
          destinationIds[index] !== idKey(to),
      )
    )
      return false;
    claimedSources.add(sourceId);
    nestedTables.add(destination);
  }
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
    if (move.kind === "empty") {
      if (
        source?.content.length !== 0 ||
        destination?.content.length !== 0 ||
        destination.pPrMark?.kind !== "ins" ||
        destination.pPrMark.info.id !== move.destinationMark
      )
        return false;
      continue;
    }
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
  if (
    geometry?.type !== "mergeCells" ||
    (geometry.moves === undefined && geometry.nestedMoves === undefined)
  )
    return table;
  const destinations = new Map(
    (geometry.moves ?? []).map((move) => [idKey(move.destination), move]),
  );
  const sources = new Map(
    (geometry.moves ?? []).map((move) => [idKey(move.source), move]),
  );
  const nested = new Map(
    (geometry.nestedMoves ?? []).flatMap((move) =>
      move.rows.map((id) => [id, move] as const),
    ),
  );
  const emptyOwner =
    geometry.emptyOwner === undefined ? undefined : idKey(geometry.emptyOwner);
  const rows = table.rows.map((row) =>
    Object.assign({}, row, {
      cells: row.cells.map((cell) =>
        Object.assign({}, cell, {
          content: cell.content.flatMap((block): BlockContent[] => {
            if (block.type === "table") {
              const move = block.rows
                .map((nestedRow) => nestedRow.structuralChange)
                .find(
                  (change) =>
                    change?.type === "tableRowInsertion" &&
                    nested.has(change.info.id),
                );
              const record =
                move === undefined ? undefined : nested.get(move.info.id);
              if (record === undefined) return [block];
              if (reject) return [];
              const placed = withNestedParagraphIds(
                block,
                new Map(
                  record.paragraphs.map(([source, destination]) => [
                    idKey(destination),
                    source,
                  ]),
                ),
              );
              const stamps = new Set(record.rows);
              placed.rows = placed.rows.map((child) => {
                const change = child.structuralChange;
                if (change === undefined || !stamps.has(change.info.id))
                  return child;
                const { structuralChange: _owned, ...accepted } = child;
                return accepted;
              });
              return [placed];
            }
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
            let revision: number | undefined;
            let range: number | undefined;
            if (move.kind !== "empty") {
              revision = reject ? move.fromRevision : move.toRevision;
              range = reject ? move.fromRange : move.toRange;
            }
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
            const restored =
              reject || move.kind === "empty" || move.controls === undefined
                ? content
                : withControlIds(
                    content,
                    new Map(
                      move.controls.map(([old, temporary]) => [temporary, old]),
                    ),
                  );
            const paragraph = {
              ...block,
              paraId: move.source,
              content: restored,
            };
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
