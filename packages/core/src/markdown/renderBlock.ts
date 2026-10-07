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

import type { BlockContent, DocxPackage } from "../types/document";
import { projectAcceptedBlocks } from "../internal/acceptedBlockProjection";
import { pushWarning } from "./internals";
import { renderParagraphBlock } from "./renderParagraph";
import { renderTable } from "./renderTable";
import type { RenderContext } from "./types";

export function renderBlocks(
  ctx: RenderContext,
  pkg: DocxPackage | undefined,
  blocks: BlockContent[],
): string {
  const outerListIndentWidths = ctx.listIndentWidths;
  ctx.listIndentWidths = [];
  const out: string[] = [];
  let prevWasListItem = false;

  const projection =
    ctx.opts.trackedChanges === "clean" ? projectAcceptedBlocks(blocks, pkg) : null;
  if (projection?.completeness === "partial") {
    pushWarning(ctx, "Some tracked structural changes could not be resolved.");
  }
  const ordered = projection?.blocks ?? blocks;

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
        ctx.listIndentWidths = [];
        const md = renderTable(ctx, pkg, block);
        ctx.listIndentWidths = [];
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
        ctx.listIndentWidths = [];
        const nested = renderBlocks(ctx, pkg, block.content);
        if (nested) {
          if (out.length) {
            out.push("");
          }
          out.push(nested);
        }
        ctx.listIndentWidths = [];
        prevWasListItem = false;
        break;
      }
      // Markup folio keeps opaquely, with no text it can claim to render.
      case "preservedBlock": {
        const previousLength = out.length;
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
        if (out.length !== previousLength) {
          ctx.listIndentWidths = [];
          prevWasListItem = false;
        }
        break;
      }
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

  ctx.listIndentWidths = outerListIndentWidths;
  return out.join("\n");
}
