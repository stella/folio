import { isDeepStrictEqual } from "node:util";

import { documentStories, findStoryBody, storyBody } from "@stll/docx-core/ops";

import type {
  BlockContent,
  Document,
  Paragraph,
} from "../../packages/docx-core/src/model/document";

const onlyFields = (value: object, allowed: readonly string[]) =>
  Object.keys(value).every((key) => allowed.includes(key));

/** No authored property, range, content or unknown metadata can become a filler. */
const bareParagraph = (
  block: BlockContent,
): block is Paragraph & { paraId: string } =>
  block.type === "paragraph" &&
  typeof block.paraId === "string" &&
  block.content.length === 0 &&
  onlyFields(block, ["type", "paraId", "content", "pPrMark"]);

const bareAddedParagraph = (block: BlockContent) =>
  bareParagraph(block) &&
  block.pPrMark !== undefined &&
  (block.pPrMark.kind === "ins" || block.pPrMark.kind === "moveTo") &&
  onlyFields(block.pPrMark, ["kind", "info"]) &&
  onlyFields(block.pPrMark.info, ["id", "author", "date"]);

const withoutField = (value: object, field: string) =>
  Object.fromEntries(Object.entries(value).filter(([key]) => key !== field));

type RejectedParagraphFillersOptions = {
  source: Document;
  actual: Document;
  expected: Document;
};

/**
 * An imported local container made entirely of bare added marks has no
 * surviving source mark after rejection. Engines may reuse different source
 * IDs for its one required empty filler. Prove both fillers at the same source
 * owner before returning any equivalence; all other IDs and data stay exact.
 */
export const rejectedParagraphFillerIdentities = ({
  source: sourceDocument,
  actual: actualDocument,
  expected: expectedDocument,
}: RejectedParagraphFillersOptions) => {
  const pairs: { actualId: string; expectedId: string }[] = [];
  type BlockLists = {
    source: readonly BlockContent[];
    actual: readonly BlockContent[];
    expected: readonly BlockContent[];
  };
  const walk = ({ source, actual, expected }: BlockLists): void => {
    if (source.length >= 2 && source.every(bareAddedParagraph)) {
      const left = actual.at(0);
      const right = expected.at(0);
      const sourceIds = new Set(
        source.map((block) =>
          block.type === "paragraph" ? block.paraId : undefined,
        ),
      );
      if (
        actual.length === 1 &&
        expected.length === 1 &&
        left !== undefined &&
        right !== undefined &&
        bareParagraph(left) &&
        bareParagraph(right) &&
        left.pPrMark === undefined &&
        right.pPrMark === undefined &&
        sourceIds.size === source.length &&
        sourceIds.has(left.paraId) &&
        sourceIds.has(right.paraId) &&
        isDeepStrictEqual(
          withoutField(left, "paraId"),
          withoutField(right, "paraId"),
        )
      )
        pairs.push({ actualId: left.paraId, expectedId: right.paraId });
      return;
    }
    // A preceding resolution may shift a structural path. That uncertainty
    // proves no filler equivalence, rather than matching a nearby container.
    if (source.length !== actual.length || source.length !== expected.length)
      return;
    for (const [index, block] of source.entries()) {
      const left = actual.at(index);
      const right = expected.at(index);
      if (left === undefined || right === undefined) continue;
      if (
        (block.type === "blockSdt" || block.type === "blockCustomXml") &&
        (left.type === "blockSdt" || left.type === "blockCustomXml") &&
        (right.type === "blockSdt" || right.type === "blockCustomXml") &&
        isDeepStrictEqual(
          withoutField(block, "content"),
          withoutField(left, "content"),
        ) &&
        isDeepStrictEqual(
          withoutField(block, "content"),
          withoutField(right, "content"),
        )
      )
        walk({
          source: block.content,
          actual: left.content,
          expected: right.content,
        });
      if (
        block.type !== "table" ||
        left.type !== "table" ||
        right.type !== "table"
      )
        continue;
      if (
        !isDeepStrictEqual(
          withoutField(block, "rows"),
          withoutField(left, "rows"),
        ) ||
        !isDeepStrictEqual(
          withoutField(block, "rows"),
          withoutField(right, "rows"),
        ) ||
        block.rows.length !== left.rows.length ||
        block.rows.length !== right.rows.length
      )
        continue;
      for (const [rowIndex, row] of block.rows.entries()) {
        const leftRow = left.rows.at(rowIndex);
        const rightRow = right.rows.at(rowIndex);
        if (
          leftRow === undefined ||
          rightRow === undefined ||
          !isDeepStrictEqual(
            withoutField(row, "cells"),
            withoutField(leftRow, "cells"),
          ) ||
          !isDeepStrictEqual(
            withoutField(row, "cells"),
            withoutField(rightRow, "cells"),
          ) ||
          row.cells.length !== leftRow.cells.length ||
          row.cells.length !== rightRow.cells.length
        )
          continue;
        for (const [cellIndex, cell] of row.cells.entries()) {
          const leftCell = leftRow.cells.at(cellIndex);
          const rightCell = rightRow.cells.at(cellIndex);
          if (
            leftCell === undefined ||
            rightCell === undefined ||
            !isDeepStrictEqual(
              withoutField(cell, "content"),
              withoutField(leftCell, "content"),
            ) ||
            !isDeepStrictEqual(
              withoutField(cell, "content"),
              withoutField(rightCell, "content"),
            )
          )
            continue;
          walk({
            source: cell.content,
            actual: leftCell.content,
            expected: rightCell.content,
          });
        }
      }
    }
  };
  for (const story of documentStories(sourceDocument)) {
    const left = findStoryBody(actualDocument, story);
    const right = findStoryBody(expectedDocument, story);
    if (left !== undefined && right !== undefined)
      walk({
        source: storyBody(sourceDocument, story).content,
        actual: left.content,
        expected: right.content,
      });
  }
  return pairs;
};
