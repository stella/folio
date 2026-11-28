import type { Document, ParagraphContent, Table } from "../model/document";
import type { TableGeometryNativeMove } from "../model/formatting";
import { storyBody, storyParagraphs } from "./blocks";
import { DOCUMENT_OP_TYPES, type DocumentOp, type OpStory } from "./types";

type MoveHalf = {
  kind: "moveFrom" | "moveTo";
  name: string;
  range: number;
  revision: number;
};

/** Read consecutive native range/payload/range triples, leaving other content in order. */
export const nativeMovesInParagraph = (
  content: readonly ParagraphContent[],
): { content: ParagraphContent[]; halves: MoveHalf[] } | undefined => {
  const plain: ParagraphContent[] = [];
  const halves: MoveHalf[] = [];
  let index = 0;
  while (index < content.length) {
    const start = content.at(index);
    if (start === undefined) return undefined;
    switch (start.type) {
      case "moveFromRangeStart":
      case "moveToRangeStart": {
        const tracked = content.at(index + 1);
        const end = content.at(index + 2);
        const kind =
          start.type === "moveFromRangeStart" ? "moveFrom" : "moveTo";
        if (
          tracked?.type !== kind ||
          end?.type !==
            (kind === "moveFrom" ? "moveFromRangeEnd" : "moveToRangeEnd") ||
          end.id !== start.id ||
          start.name.length === 0 ||
          start.copyOriginAncestors !== undefined ||
          end.copyOriginAncestors !== undefined ||
          tracked.resolutionJoins !== undefined
        )
          return undefined;
        halves.push({
          kind,
          name: start.name,
          range: start.id,
          revision: tracked.info.id,
        });
        plain.push(...tracked.content);
        index += 3;
        break;
      }
      case "moveFrom":
      case "moveTo":
      case "moveFromRangeEnd":
      case "moveToRangeEnd":
        return undefined;
      default:
        plain.push(start);
        index += 1;
    }
  }
  return { content: plain, halves };
};

export const pairedNativeCellMoves = (
  halves: readonly MoveHalf[],
): TableGeometryNativeMove[] | undefined => {
  const byName = new Map<string, { moveFrom?: MoveHalf; moveTo?: MoveHalf }>();
  const ranges = new Set<number>();
  const revisions = new Set<number>();
  for (const half of halves) {
    const group = byName.get(half.name) ?? {};
    if (
      group[half.kind] !== undefined ||
      ranges.has(half.range) ||
      revisions.has(half.revision)
    )
      return undefined;
    group[half.kind] = half;
    byName.set(half.name, group);
    ranges.add(half.range);
    revisions.add(half.revision);
  }
  const moves: TableGeometryNativeMove[] = [];
  for (const [name, { moveFrom, moveTo }] of byName) {
    if (moveFrom === undefined || moveTo === undefined) return undefined;
    moves.push({
      name,
      fromRange: moveFrom.range,
      toRange: moveTo.range,
      fromRevision: moveFrom.revision,
      toRevision: moveTo.revision,
    });
  }
  return moves;
};

/** After the table merge, both native wrapper IDs may be picked directly. */
export const selectedNativeTableMoves = (
  tables: readonly Table[],
  revisionIds: ReadonlySet<number>,
): TableGeometryNativeMove[] => {
  const candidates = new Map<string, MoveHalf[]>();
  for (const table of tables) {
    if (table.formatting?.gridChange?.geometryChange !== undefined) continue;
    for (const { cells } of table.rows) {
      for (const { content } of cells) {
        for (const block of content) {
          if (block.type !== "paragraph" || block.pPrMark !== undefined)
            continue;
          if (
            !block.content.some(
              (node) =>
                (node.type === "moveFrom" || node.type === "moveTo") &&
                revisionIds.has(node.info.id),
            )
          )
            continue;
          const native = nativeMovesInParagraph(block.content);
          for (const half of native?.halves ?? []) {
            if (!revisionIds.has(half.revision)) continue;
            const matching = candidates.get(half.name);
            if (matching === undefined) candidates.set(half.name, [half]);
            else matching.push(half);
          }
        }
      }
    }
  }
  return [...candidates.values()].flatMap(
    (halves) => pairedNativeCellMoves(halves) ?? [],
  );
};

type NativeRangeReviewOptions = {
  document: Document;
  story: OpStory;
  moves: readonly TableGeometryNativeMove[];
};

/** Remove boundaries only beside the physical payload selected for review. */
export const settledNativeMoveRangeEdits = ({
  document,
  story,
  moves,
}: NativeRangeReviewOptions): DocumentOp[] => {
  if (moves.length === 0) return [];
  const from = new Map(moves.map((move) => [move.fromRevision, move]));
  const to = new Map(moves.map((move) => [move.toRevision, move]));
  const edits: DocumentOp[] = [];
  for (const { paragraph } of storyParagraphs(storyBody(document, story))) {
    if (paragraph.paraId === undefined) continue;
    const parsed = nativeMovesInParagraph(paragraph.content);
    const retiring = new Set<string>();
    for (const half of parsed?.halves ?? []) {
      const move = (half.kind === "moveFrom" ? from : to).get(half.revision);
      if (
        move?.name === half.name &&
        half.range ===
          (half.kind === "moveFrom" ? move.fromRange : move.toRange)
      )
        retiring.add(`${half.kind}:${half.range}`);
    }
    if (retiring.size === 0) continue;
    const content = paragraph.content.filter((node) => {
      switch (node.type) {
        case "moveFromRangeStart":
        case "moveFromRangeEnd":
          return !retiring.has(`moveFrom:${node.id}`);
        case "moveToRangeStart":
        case "moveToRangeEnd":
          return !retiring.has(`moveTo:${node.id}`);
        default:
          return true;
      }
    });
    edits.push({
      type: DOCUMENT_OP_TYPES.REPLACE_INLINE,
      story,
      blockId: paragraph.paraId,
      expected: paragraph.content,
      content,
    });
  }
  return edits;
};
