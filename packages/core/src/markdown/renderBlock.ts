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
  TrackedRunContent,
} from "../types/document";
import { cloneParagraphWithoutPropertySource } from "../docx/paragraphPropertySource";
import { renderParagraphBlock } from "./renderParagraph";
import { renderTable } from "./renderTable";
import type { RenderContext } from "./types";

/**
 * In `trackedChanges: "clean"` mode every change is accepted. A paragraph whose
 * end-of-paragraph mark is a pending deletion (`pPrMark.kind === "del"`) loses
 * its break on accept and merges with the following paragraph. Word's join
 * keeps the FIRST paragraph's properties (style, list) and drops the resolved
 * mark; the surviving break is the next paragraph's, so a run of consecutive
 * deletions collapses into one paragraph. A first paragraph with nothing left
 * once its deletions are accepted (a whole deleted paragraph) contributes only
 * the join, so the NEXT paragraph keeps its own properties, as the editor's
 * accept does. A non-paragraph next block (table, SDT) is structurally
 * incompatible and stays unmerged, matching the editor's accept-change join
 * guard (`commands/comments.ts`).
 */
function mergeAcceptedParagraphBreaks(blocks: BlockContent[]): BlockContent[] {
  const merged: BlockContent[] = [];
  for (const block of blocks) {
    const prev = merged.at(-1);
    if (prev?.type === "paragraph" && prev.pPrMark?.kind === "del" && block.type === "paragraph") {
      // Drop the resolved deletion mark; inherit the next paragraph's mark so a
      // chain keeps merging.
      const next = block.pPrMark;
      const content = [...prev.content, ...block.content];
      const formattingOwner = holdsNothingOnAccept(prev) ? block : prev;
      const joined = cloneParagraphWithoutPropertySource(formattingOwner, {
        content,
        ...(next ? { pPrMark: next } : {}),
      });
      if (!next) {
        Reflect.deleteProperty(joined, "pPrMark");
      }
      merged[merged.length - 1] = joined;
      continue;
    }
    merged.push(block);
  }
  return merged;
}

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
