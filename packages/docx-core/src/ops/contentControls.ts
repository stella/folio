/**
 * Set one content control's properties, block-level or inline, wherever it
 * sits in its story. This is the one way a control's value changes: the
 * dropdown, combo box, date and checkbox widgets set `w:lastValue`,
 * `w:fullDate` or `w14:checked` here, and edit the shown text with the text
 * operations in the same batch.
 */

import { Result } from "better-result";

import type {
  BlockContent,
  BlockSdt,
  Document,
  Paragraph,
  SdtProperties,
} from "../model/document";
import { withBodyContent } from "./blocks";
import type { DocumentEdit } from "./edits";
import { equalForStaleness, structurallyEqual } from "./equality";
import { idKey } from "./ids";
import {
  asParagraphContent,
  childNodes,
  type InlineNode,
  rebuildNode,
} from "./leaves";
import { cloneModel } from "./modelClone";
import {
  DOCUMENT_OP_REFUSAL_REASONS,
  DocumentOpRefusal,
  type DocumentOpRefusalReason,
} from "./refusal";
import { replaceStoryBody, storyBody } from "./stories";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type SetContentControlOp,
} from "./types";

type Applied = Result<DocumentEdit, DocumentOpRefusal>;

/** What one walk found: each matching control's properties, and the paragraphs it touched. */
type Found = {
  controls: SdtProperties[];
  paragraphs: Paragraph[];
};

const refuse = (
  op: SetContentControlOp,
  reason: DocumentOpRefusalReason,
  message: string,
): Applied =>
  Result.err(new DocumentOpRefusal({ opType: op.type, reason, message }));

const matches = (properties: SdtProperties, controlId: number): boolean =>
  properties.id === controlId;

/** Every paragraph inside a list of blocks, nested containers included. */
const paragraphsIn = (blocks: readonly BlockContent[]): Paragraph[] =>
  blocks.flatMap((block): Paragraph[] => {
    switch (block.type) {
      case "paragraph":
        return [block];
      case "table":
        return block.rows.flatMap((row) =>
          row.cells.flatMap((cell) => paragraphsIn(cell.content)),
        );
      case "blockSdt":
      case "blockCustomXml":
        return paragraphsIn(block.content);
      case "preservedBlock":
      case "bookmarkStart":
      case "bookmarkEnd":
      case "commentRangeStart":
      case "commentRangeEnd":
      case "moveFromRangeStart":
      case "moveFromRangeEnd":
      case "moveToRangeStart":
      case "moveToRangeEnd":
        return [];
      default: {
        const unreachable: never = block;
        return unreachable;
      }
    }
  });

/** The innermost block-level control holding a paragraph, and every paragraph it holds. */
export const blockControlHolding = (
  blocks: readonly BlockContent[],
  blockId: string,
): { control: BlockSdt; paragraphs: Paragraph[] } | undefined => {
  for (const block of blocks) {
    let nested: readonly BlockContent[][] = [];
    if (block.type === "table")
      nested = block.rows.flatMap((row) =>
        row.cells.map((cell) => cell.content),
      );
    else if (block.type === "blockSdt" || block.type === "blockCustomXml")
      nested = [block.content];
    for (const content of nested) {
      const inner = blockControlHolding(content, blockId);
      if (inner !== undefined) return inner;
    }
    if (block.type !== "blockSdt") continue;
    const paragraphs = paragraphsIn(block.content);
    if (
      paragraphs.some(
        ({ paraId }) =>
          paraId !== undefined && idKey(paraId) === idKey(blockId),
      )
    )
      return { control: block, paragraphs };
  }
  return undefined;
};

type Rewrite = {
  controlId: number;
  properties: SdtProperties;
  found: Found;
};

const rewriteInline = (
  nodes: readonly InlineNode[],
  rewrite: Rewrite,
): readonly InlineNode[] => {
  let changed = false;
  const next = nodes.map((node) => {
    let current: InlineNode = node;
    if (
      current.type === "inlineSdt" &&
      matches(current.properties, rewrite.controlId)
    ) {
      rewrite.found.controls.push(current.properties);
      current = { ...current, properties: cloneModel(rewrite.properties) };
    }
    const children = childNodes(current);
    if (children !== undefined) {
      const rewritten = rewriteInline(children, rewrite);
      if (rewritten !== children) current = rebuildNode(current, rewritten);
    }
    if (current !== node) changed = true;
    return current;
  });
  return changed ? next : nodes;
};

const rewriteParagraph = (
  paragraph: Paragraph,
  rewrite: Rewrite,
): Paragraph => {
  const before = rewrite.found.controls.length;
  const content = rewriteInline(paragraph.content, rewrite);
  if (rewrite.found.controls.length === before) return paragraph;
  rewrite.found.paragraphs.push(paragraph);
  return { ...paragraph, content: asParagraphContent(content) };
};

const rewriteBlocks = (
  blocks: readonly BlockContent[],
  rewrite: Rewrite,
): readonly BlockContent[] => {
  let changed = false;
  const next = blocks.map((block): BlockContent => {
    let current: BlockContent = block;
    switch (block.type) {
      case "paragraph":
        current = rewriteParagraph(block, rewrite);
        break;
      case "table": {
        const rows = block.rows.map((row) => {
          const cells = row.cells.map((cell) => {
            const content = rewriteBlocks(cell.content, rewrite);
            return content === cell.content
              ? cell
              : { ...cell, content: [...content] };
          });
          return cells.every((cell, index) => cell === row.cells[index])
            ? row
            : { ...row, cells };
        });
        if (rows.some((row, index) => row !== block.rows[index]))
          current = { ...block, rows };
        break;
      }
      case "blockSdt": {
        let properties = block.properties;
        if (matches(block.properties, rewrite.controlId)) {
          rewrite.found.controls.push(block.properties);
          rewrite.found.paragraphs.push(...paragraphsIn(block.content));
          properties = cloneModel(rewrite.properties);
        }
        const content = rewriteBlocks(block.content, rewrite);
        if (properties !== block.properties || content !== block.content)
          current = { ...block, properties, content: [...content] };
        break;
      }
      case "blockCustomXml": {
        const content = rewriteBlocks(block.content, rewrite);
        if (content !== block.content)
          current = { ...block, content: [...content] };
        break;
      }
      case "preservedBlock":
      case "bookmarkStart":
      case "bookmarkEnd":
      case "commentRangeStart":
      case "commentRangeEnd":
      case "moveFromRangeStart":
      case "moveFromRangeEnd":
      case "moveToRangeStart":
      case "moveToRangeEnd":
        break;
      default: {
        const unreachable: never = block;
        return unreachable;
      }
    }
    if (current !== block) changed = true;
    return current;
  });
  return changed ? next : blocks;
};

/** Apply a {@link SetContentControlOp}; its inverse swaps `expected` and `properties`. */
export const setContentControl = (
  document: Document,
  op: SetContentControlOp,
): Applied => {
  if (
    op.properties.id !== op.controlId ||
    op.expected.id !== op.controlId ||
    op.properties.sdtType !== op.expected.sdtType
  )
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.INVALID_OPERATION,
      "A content control keeps its id and its type.",
    );
  const body = storyBody(document, op.story);
  const rewrite: Rewrite = {
    controlId: op.controlId,
    properties: op.properties,
    found: { controls: [], paragraphs: [] },
  };
  const content = rewriteBlocks(body.content, rewrite);
  const [control, ...others] = rewrite.found.controls;
  if (control === undefined)
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
      `No content control has id ${String(op.controlId)}.`,
    );
  if (others.length > 0)
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.DUPLICATE_RECORD_ID,
      `More than one content control has id ${String(op.controlId)}.`,
    );
  const blockIds = rewrite.found.paragraphs.flatMap(({ paraId }) =>
    paraId === undefined ? [] : [paraId],
  );
  if (
    blockIds.length !== op.blockIds.length ||
    blockIds.some((id, index) => idKey(id) !== idKey(op.blockIds[index] ?? ""))
  )
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STALE,
      `Content control ${String(op.controlId)} holds other paragraphs than expected.`,
    );
  if (!equalForStaleness(control, op.expected))
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STALE,
      `Content control ${String(op.controlId)} has other properties than expected.`,
    );
  if (structurallyEqual(control, op.properties))
    return Result.ok({
      document,
      inverse: [],
      touched: { modified: [], inserted: [], removed: [] },
    });
  const inverse: DocumentOp = {
    type: DOCUMENT_OP_TYPES.SET_CONTENT_CONTROL,
    story: op.story,
    controlId: op.controlId,
    blockIds: [...op.blockIds],
    expected: cloneModel(op.properties),
    properties: cloneModel(control),
  };
  return Result.ok({
    document: replaceStoryBody({
      document,
      story: op.story,
      // The body's section view holds its blocks too, so it follows them.
      body: withBodyContent(body, [...content]),
    }),
    inverse: [inverse],
    touched: {
      modified: blockIds,
      inserted: [],
      removed: [],
    },
  });
};
