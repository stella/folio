/**
 * Block-level dispatcher. Walks `BlockContent[]` and joins the rendered
 * markdown for each block, suppressing redundant blank lines between list items
 * of the same list. Ported from eigenpal/docx-editor PR #595.
 */

import { panic } from "better-result";
import { ALT_CHUNK_READER_DIAGNOSTIC, isAltChunkMarkup } from "../docx/altChunk";
import {
  OPAQUE_REVISION_CARRIER_READER_DIAGNOSTIC,
  isOpaqueNestedRowMarkup,
  opaqueRevisionCarrierName,
} from "../docx/opaqueCarrier";

import type {
  BlockContent,
  DocxPackage,
  Paragraph,
  ParagraphContent,
  Table,
  TrackedRunContent,
} from "../types/document";
import { cloneParagraphWithoutPropertySource } from "../docx/paragraphPropertySource";
import { renderParagraphBlock } from "./renderParagraph";
import { renderTable } from "./renderTable";
import type { RenderContext } from "./types";

/**
 * In `trackedChanges: "clean"` mode every change is accepted, as the editor's
 * accept-all resolves the story (`internal/wholeStoryRevisionResolution.ts`).
 *
 * A paragraph whose mark is pending deletion loses its break on accept and
 * runs on into the paragraph after it, which is the paragraph left: its mark
 * ends the joined text, and a paragraph's properties (style, list) live on its
 * mark. A run of consecutive deletions collapses into that last paragraph.
 *
 * With a table right after it, its words run on into the table's first cell
 * and lead that cell's first paragraph (`prosemirror/tableRunIn.ts`), unless
 * it ends a section, keeps nothing once accepted, or that cell is pending.
 * Otherwise a paragraph that cannot join keeps its place, or goes when it
 * keeps nothing and something follows it.
 */
function mergeAcceptedParagraphBreaks(blocks: BlockContent[]): BlockContent[] {
  const reversed: BlockContent[] = [];
  // The paragraph the ones before it run into: the one right after, when it is one.
  let survivor: Paragraph | null = null;
  // The block right after is a table, or a paragraph that ran into one.
  let tableFollows = false;
  let followed = false;
  const flush = (): void => {
    if (survivor) reversed.push(survivor);
    survivor = null;
  };
  for (let index = blocks.length - 1; index >= 0; index--) {
    const block = blocks[index];
    if (!block) continue;
    if (block.type !== "paragraph") {
      flush();
      reversed.push(block);
      tableFollows = block.type === "table";
      followed ||= tableFollows;
      continue;
    }
    const goes = block.pPrMark?.kind === "del" || block.pPrMark?.kind === "moveFrom";
    if (goes && survivor) {
      survivor = cloneParagraphWithoutPropertySource(survivor, {
        content: [...block.content, ...survivor.content],
      });
      continue;
    }
    const table = reversed.at(-1);
    if (goes && tableFollows && table?.type === "table") {
      const ranIn = tableWithParagraphRunIn(table, block);
      if (ranIn) {
        reversed[reversed.length - 1] = ranIn;
        continue;
      }
    }
    flush();
    tableFollows = false;
    if (goes && holdsNothingOnAccept(block) && followed && index + reversed.length > 0) {
      continue;
    }
    survivor = block;
    followed = true;
  }
  flush();
  return reversed.reverse();
}

/**
 * Whether the cell a paragraph's words would run into is itself pending (its
 * row or the cell, through nested tables' first cells), as
 * `prosemirror/tableRunIn.ts` reads it.
 */
const runInCellPending = (table: Table): boolean => {
  const row = table.rows[0];
  const cell = row?.cells[0];
  if (!row || !cell) return false;
  if (row.structuralChange !== undefined || cell.structuralChange !== undefined) return true;
  const first = cell.content[0];
  return first?.type === "table" && runInCellPending(first);
};

/**
 * `table` with `paragraph`'s words leading its first cell's first paragraph,
 * descending into nested tables, or null when the paragraph keeps its place.
 */
const tableWithParagraphRunIn = (table: Table, paragraph: Paragraph): Table | null => {
  if (paragraph.sectionProperties !== undefined || holdsNothingOnAccept(paragraph)) return null;
  if (runInCellPending(table)) return null;
  const runIn = (current: Table): Table | null => {
    const [row, ...rows] = current.rows;
    const [cell, ...cells] = row?.cells ?? [];
    const [first, ...rest] = cell?.content ?? [];
    if (!row || !cell || !first) return null;
    let led: BlockContent | null = null;
    if (first.type === "paragraph") {
      led = cloneParagraphWithoutPropertySource(first, {
        content: [...paragraph.content, ...first.content],
      });
    } else if (first.type === "table") {
      led = runIn(first);
    }
    if (!led) return null;
    return {
      ...current,
      rows: [{ ...row, cells: [{ ...cell, content: [led, ...rest] }, ...cells] }, ...rows],
    };
  };
  return runIn(table);
};

/** Paragraph content that shows nothing: range boundaries and anchors. */
const ZERO_WIDTH_CONTENT: ReadonlySet<string> = new Set([
  "bookmarkStart",
  "bookmarkEnd",
  "commentRangeStart",
  "commentRangeEnd",
  "commentReference",
  "moveFromRangeStart",
  "moveFromRangeEnd",
  "moveToRangeStart",
  "moveToRangeEnd",
]);

/** Whether an item shows nothing once every change is accepted. */
const showsNothingOnAccept = (item: ParagraphContent | TrackedRunContent): boolean => {
  if (item.type === "deletion" || item.type === "moveFrom" || ZERO_WIDTH_CONTENT.has(item.type)) {
    return true;
  }
  // An insertion whose text was deleted again (`w:ins > w:del`).
  if (item.type === "insertion" || item.type === "moveTo") {
    return item.content.every(showsNothingOnAccept);
  }
  return (
    item.type === "run" && item.content.every((content) => content.type === "renderedPageBreak")
  );
};

/** Whether a paragraph keeps no visible content once every change is accepted. */
const holdsNothingOnAccept = (paragraph: Paragraph): boolean =>
  paragraph.content.every(showsNothingOnAccept);

export function renderBlocks(
  ctx: RenderContext,
  pkg: DocxPackage | undefined,
  blocks: BlockContent[],
): string {
  const out: string[] = [];
  let prevWasListItem = false;

  const ordered =
    ctx.opts.trackedChanges === "clean" ? mergeAcceptedParagraphBreaks(blocks) : blocks;

  for (const block of ordered) {
    switch (block.type) {
      case "paragraph": {
        // A hidden-marker (`w:vanish`) list paragraph renders as plain prose,
        // and a numbered heading as a heading, so neither joins a run of list
        // items (neither may suppress the blank line around one).
        const { markdown: md, isListItem } = renderParagraphBlock(ctx, pkg, block);
        if (!md) {
          prevWasListItem = false;
          continue;
        }
        if (isListItem && prevWasListItem) {
          out.push(md);
        } else if (out.length) {
          out.push("", md);
        } else {
          out.push(md);
        }
        prevWasListItem = isListItem;
        break;
      }
      case "table": {
        const md = renderTable(ctx, pkg, block);
        if (md) {
          if (out.length) {
            out.push("");
          }
          out.push(md);
        }
        prevWasListItem = false;
        break;
      }
      case "blockSdt":
      case "blockCustomXml": {
        const nested = renderBlocks(ctx, pkg, block.content);
        if (nested) {
          if (out.length) {
            out.push("");
          }
          out.push(nested);
        }
        break;
      }
      // Markup folio keeps opaquely, with no text it can claim to render.
      case "preservedBlock":
        if (isAltChunkMarkup(block.xml)) {
          out.push(
            block.readerText === undefined
              ? ALT_CHUNK_READER_DIAGNOSTIC
              : `${ALT_CHUNK_READER_DIAGNOSTIC}\n\n${block.readerText}`,
          );
        } else if (opaqueRevisionCarrierName(block.xml) !== undefined) {
          out.push(OPAQUE_REVISION_CARRIER_READER_DIAGNOSTIC);
        } else if (isOpaqueNestedRowMarkup(block.xml)) {
          out.push("[Unsupported nested w:tr content]");
        }
        break;
      // A delimiter, with no text to render.
      case "bookmarkStart":
      case "bookmarkEnd":
        break;
      default: {
        const unsupported: never = block;
        panic(`Unsupported block content in markdown: ${JSON.stringify(unsupported)}`);
      }
    }
  }

  return out.join("\n");
}
