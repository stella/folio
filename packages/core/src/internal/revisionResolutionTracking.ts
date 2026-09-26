import { panic } from "better-result";
import type { Node as PMNode } from "prosemirror-model";

import { indexedPositionMap } from "./indexedPositionMap";
import type { resolveWholeStory } from "./wholeStoryRevisionResolution";

type Resolution = ReturnType<typeof resolveWholeStory>;
type Range = { from: number; to: number };
type ParagraphMapping = { source: number; sourceEnd: number; final: number; finalEnd: number };
type TableTracking = Resolution["tableRanges"][number] & {
  finalPosition: number | null;
  parent: TableTracking | null;
  paragraphs: readonly ParagraphMapping[];
  mapInput: (position: number, assoc: 1 | -1) => number;
};

const paragraphSizes = (table: PMNode) => {
  const sizes = new Map<number, number>();
  table.descendants((node, position) => {
    if (node.type.name === "table") {
      return false;
    }
    if (node.type.name === "paragraph") {
      sizes.set(position + 1, node.nodeSize);
      return false;
    }
    return true;
  });
  return sizes;
};

/** Map the three resolution phases to precise ranges in the finished story. */
export const finalRevisionParagraphRanges = (result: Resolution): Range[] => {
  const tableNodes = new Set(result.tableRanges.map(({ node }) => node));
  const finalTablePositions = new Map<PMNode, number>();
  result.resolved.descendants((node, position) => {
    if (tableNodes.has(node)) {
      finalTablePositions.set(node, position);
    }
    return true;
  });

  const tables: TableTracking[] = result.tableRanges
    .map((table) => {
      const finalPosition = finalTablePositions.get(table.node) ?? null;
      const sourceSizes = paragraphSizes(table.inputNode);
      const finalSizes =
        finalPosition === null ? new Map<number, number>() : paragraphSizes(table.node);
      const paragraphs: ParagraphMapping[] = table.paragraphOffsets.map(({ source, final }) => {
        const sourceSize = sourceSizes.get(source);
        const finalSize = finalSizes.get(final);
        if (sourceSize === undefined || (finalPosition !== null && finalSize === undefined)) {
          return panic("Resolved table paragraph offsets are inconsistent");
        }
        return {
          source,
          sourceEnd: source + sourceSize,
          final,
          finalEnd: final + (finalSize ?? 0),
        };
      });
      paragraphs.sort((left, right) => left.source - right.source);
      return {
        ...table,
        finalPosition,
        parent: null,
        paragraphs,
        mapInput: indexedPositionMap(table.inputMap),
      };
    })
    .sort(
      (left, right) => left.position - right.position || right.sourceNodeSize - left.sourceNodeSize,
    );
  const tableStack: TableTracking[] = [];
  for (const table of tables) {
    for (;;) {
      const enclosing = tableStack.at(-1);
      if (!enclosing || table.position < enclosing.position + enclosing.sourceNodeSize) {
        break;
      }
      tableStack.pop();
    }
    table.parent = tableStack.at(-1) ?? null;
    tableStack.push(table);
  }
  const mapInline = indexedPositionMap(result.inlineMap);
  const mapStructure = indexedPositionMap(result.structuralMap);
  const finalRanges: Range[] = [];

  const tableContaining = (range: Range): TableTracking | null => {
    let low = 0;
    let high = tables.length;
    while (low < high) {
      const middle = Math.floor((low + high) / 2);
      const candidate = tables.at(middle);
      if (candidate && candidate.position <= range.from) {
        low = middle + 1;
      } else {
        high = middle;
      }
    }
    let table = low === 0 ? null : (tables.at(low - 1) ?? null);
    while (table && range.to > table.position + table.sourceNodeSize) {
      table = table.parent;
    }
    return table;
  };

  const mapAfterInlineRange = ({ from, to }: Range): void => {
    if (to <= from) {
      return;
    }
    const owningTable = tableContaining({ from, to });
    if (owningTable) {
      if (owningTable.finalPosition === null) {
        return;
      }
      const relativeFrom = owningTable.mapInput(from - owningTable.position, 1);
      const relativeTo = owningTable.mapInput(to - owningTable.position, -1);
      let low = 0;
      let high = owningTable.paragraphs.length;
      while (low < high) {
        const middle = Math.floor((low + high) / 2);
        const candidate = owningTable.paragraphs.at(middle);
        if (candidate && candidate.sourceEnd <= relativeFrom) {
          low = middle + 1;
        } else {
          high = middle;
        }
      }
      for (let index = low; index < owningTable.paragraphs.length; index++) {
        const paragraph = owningTable.paragraphs.at(index);
        if (!paragraph || paragraph.source >= relativeTo) {
          break;
        }
        finalRanges.push({
          from: owningTable.finalPosition + paragraph.final,
          to: owningTable.finalPosition + paragraph.finalEnd,
        });
      }
      return;
    }
    const finalFrom = mapStructure(from, 1);
    const finalTo = mapStructure(to, -1);
    if (finalTo > finalFrom) {
      finalRanges.push({ from: finalFrom, to: finalTo });
    }
  };

  for (const range of result.changedRanges) {
    mapAfterInlineRange({
      from: mapInline(range.from, 1),
      to: mapInline(range.to, -1),
    });
  }
  for (const range of result.structuralRanges) {
    mapAfterInlineRange(range);
  }
  for (const table of tables) {
    if (table.finalPosition === null) {
      continue;
    }
    for (const range of table.ranges) {
      finalRanges.push({
        from: table.finalPosition + range.from,
        to: table.finalPosition + range.to,
      });
    }
  }
  return finalRanges;
};
