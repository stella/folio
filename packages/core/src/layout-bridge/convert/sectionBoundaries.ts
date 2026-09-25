/**
 * Section boundary handling for the flow block list: section start types,
 * trailing page breaks before continuous sections, and section document grids.
 */

import type { Node as PMNode } from "prosemirror-model";
import type { FlowBlock, SectionBreakBlock } from "../../layout-engine/types";
import { parseSectionBreakType } from "../../prosemirror/sectionCarrier";
import { twipsToPixels } from "./flowConversionShared";

export function getLastMapKey<K, V>(map: ReadonlyMap<K, V>): K | undefined {
  let lastKey: K | undefined;
  for (const key of map.keys()) {
    lastKey = key;
  }
  return lastKey;
}

/**
 * Translate section-owned start modes into the boundary-owned values consumed
 * by the paginator. Each flow section break carries the properties for the
 * section it ends, so its start mode belongs on the preceding boundary.
 */
export function applySectionStartsToBoundaries(
  blocks: readonly FlowBlock[],
  finalSectionStart: NonNullable<SectionBreakBlock["type"]> | undefined,
): FlowBlock[] {
  const breakIndexes: number[] = [];
  for (let index = 0; index < blocks.length; index += 1) {
    if (blocks[index]?.kind === "sectionBreak") {
      breakIndexes.push(index);
    }
  }
  if (breakIndexes.length === 0) {
    return [...blocks];
  }

  const result = [...blocks];
  for (let index = 0; index < breakIndexes.length; index += 1) {
    const boundaryIndex = breakIndexes[index];
    if (boundaryIndex === undefined) {
      continue;
    }
    const boundary = blocks[boundaryIndex];
    if (boundary?.kind !== "sectionBreak") {
      continue;
    }
    const nextBoundaryIndex = breakIndexes[index + 1];
    const nextBoundary = nextBoundaryIndex === undefined ? undefined : blocks[nextBoundaryIndex];
    const nextStart = nextBoundary?.kind === "sectionBreak" ? nextBoundary.type : finalSectionStart;
    const translated = { ...boundary };
    if (nextStart === undefined) {
      delete translated.type;
    } else {
      translated.type = nextStart;
    }
    result[boundaryIndex] = translated;
  }
  return result;
}

/**
 * A trailing page break belongs to its section-ending paragraph. When the
 * paragraph mark stays on the current page, an immediately following
 * continuous section resumes there too. Accept both the legacy generated
 * carrier and the inline atom's one-token paragraph-mark fragment; neither is
 * a physical layout boundary.
 *
 * A section break in `breaksWithoutMarker` ends at an empty w:sectPr paragraph
 * that projected no block, so a page break just before it ends an earlier
 * paragraph and still starts the next page.
 */
export function coalesceTrailingPageBreakBeforeContinuousSection(
  blocks: readonly FlowBlock[],
  splitPageBreakAndParagraphMark: boolean,
  breaksWithoutMarker: ReadonlySet<SectionBreakBlock["id"]>,
): FlowBlock[] {
  if (splitPageBreakAndParagraphMark) {
    return [...blocks];
  }

  const result: FlowBlock[] = [];
  for (let index = 0; index < blocks.length; index += 1) {
    const block = blocks[index];
    const carrier = blocks[index + 1];
    const section = blocks[index + 2];
    const isLegacyGeneratedCarrier =
      block?.kind === "pageBreak" &&
      carrier?.kind === "paragraph" &&
      carrier.runs.length === 0 &&
      block.pmStart !== undefined &&
      block.pmStart === carrier.pmStart &&
      carrier.pmStart === carrier.pmEnd;
    const isInlineParagraphMarkCarrier =
      block?.kind === "pageBreak" &&
      carrier?.kind === "paragraph" &&
      carrier.runs.length === 0 &&
      block.pmEnd !== undefined &&
      carrier.pmStart === block.pmEnd &&
      carrier.pmEnd === block.pmEnd + 1;
    const isUnresolvedInsertion = block?.kind === "pageBreak" && block.isInsertion === true;
    if (
      block?.kind === "pageBreak" &&
      !isUnresolvedInsertion &&
      carrier?.kind === "sectionBreak" &&
      carrier.type === "continuous" &&
      !breaksWithoutMarker.has(carrier.id)
    ) {
      continue;
    }
    if (
      block?.kind === "pageBreak" &&
      !isUnresolvedInsertion &&
      (isLegacyGeneratedCarrier || isInlineParagraphMarkCarrier) &&
      section?.kind === "sectionBreak" &&
      section.type === "continuous" &&
      !breaksWithoutMarker.has(section.id)
    ) {
      index += 1;
      continue;
    }
    if (block) {
      result.push(block);
    }
  }
  return result;
}

/**
 * The last section's start mode, off a document attribute the schema types as
 * `any`. `parseSectionBreakType` owns the enumeration, so a member added to
 * `ST_SectionMark` reaches the paginator instead of being dropped here.
 */
export function readFinalSectionStart(
  doc: PMNode,
): NonNullable<SectionBreakBlock["type"]> | undefined {
  return parseSectionBreakType(doc.attrs["_finalSectionStart"]) ?? undefined;
}

type SectionDocumentGridOptions = {
  finalLinePitchTwips: number | undefined;
  tableCellLinePitch: "sectionGrid" | undefined;
};

export function applySectionDocumentGrid(
  blocks: FlowBlock[],
  { finalLinePitchTwips, tableCellLinePitch }: SectionDocumentGridOptions,
): FlowBlock[] {
  const result = [...blocks];
  let sectionStart = 0;

  const stampSection = (end: number, linePitchTwips: number | undefined): void => {
    if (linePitchTwips === undefined || linePitchTwips <= 0) {
      return;
    }
    const linePitch = twipsToPixels(linePitchTwips);
    const stampBlock = (block: FlowBlock): FlowBlock => {
      if (block.kind === "paragraph") {
        return {
          ...block,
          attrs: { ...block.attrs, documentGridLinePitch: linePitch },
        };
      }
      if (block.kind !== "table" || tableCellLinePitch !== "sectionGrid") {
        return block;
      }
      return {
        ...block,
        rows: block.rows.map((row) => ({
          ...row,
          cells: row.cells.map((cell) => ({
            ...cell,
            blocks: cell.blocks.map(stampBlock),
          })),
        })),
      };
    };
    for (let index = sectionStart; index < end; index += 1) {
      const block = result[index];
      if (!block) {
        continue;
      }
      result[index] = stampBlock(block);
    }
  };

  for (let index = 0; index < blocks.length; index += 1) {
    const block = blocks[index];
    if (block?.kind !== "sectionBreak") {
      continue;
    }
    stampSection(index, block.documentGridLinePitchTwips);
    sectionStart = index + 1;
  }
  stampSection(blocks.length, finalLinePitchTwips);
  return result;
}
