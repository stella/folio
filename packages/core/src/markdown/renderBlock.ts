/**
 * Block-level dispatcher. Walks `BlockContent[]` and joins the rendered
 * markdown for each block, suppressing redundant blank lines between list items
 * of the same list. Ported from eigenpal/docx-editor PR #595.
 */

import { panic } from "better-result";

import {
  ALT_CHUNK_READER_DIAGNOSTIC,
  isAltChunkMarkup,
} from "../docx/altChunk";
import {
  isOpaqueNestedRowMarkup,
  OPAQUE_REVISION_CARRIER_READER_DIAGNOSTIC,
  opaqueRevisionCarrierName,
} from "../docx/opaqueCarrier";
import { projectAcceptedBlocks } from "../internal/acceptedBlockProjection";
import type { BlockContent, DocxPackage } from "../types/document";
import { pushWarning } from "./internals";
import {
  type RenderedParagraph,
  renderParagraphBlock,
} from "./renderParagraph";
import { renderTable } from "./renderTable";
import type { RenderContext } from "./types";

type RenderedBlock = RenderedParagraph | { type: "delimiter" };

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
    ctx.opts.trackedChanges === "clean"
      ? projectAcceptedBlocks(blocks, pkg)
      : null;
  if (projection?.completeness === "partial") {
    pushWarning(ctx, "Some tracked structural changes could not be resolved.");
  }
  const ordered = projection?.blocks ?? blocks;

  const renderBlock = (block: BlockContent): RenderedBlock => {
    switch (block.type) {
      case "paragraph": {
        // A hidden-marker (`w:vanish`) list paragraph renders as plain prose,
        // and a numbered heading as a heading, so neither joins a run of list
        // items (neither may suppress the blank line around one).
        return renderParagraphBlock(ctx, pkg, block);
      }
      case "table": {
        ctx.listIndentWidths = [];
        const md = renderTable(ctx, pkg, block);
        ctx.listIndentWidths = [];
        return { type: "paragraph", markdown: md };
      }
      case "blockSdt":
      case "blockCustomXml": {
        ctx.listIndentWidths = [];
        const nested = renderBlocks(ctx, pkg, block.content);
        ctx.listIndentWidths = [];
        return { type: "paragraph", markdown: nested };
      }
      // Markup folio keeps opaquely, with no text it can claim to render.
      case "preservedBlock": {
        let markdown: string;
        if (isAltChunkMarkup(block.xml)) {
          markdown =
            block.readerText === undefined
              ? ALT_CHUNK_READER_DIAGNOSTIC
              : `${ALT_CHUNK_READER_DIAGNOSTIC}\n\n${block.readerText}`;
        } else if (opaqueRevisionCarrierName(block.xml) !== undefined) {
          markdown = OPAQUE_REVISION_CARRIER_READER_DIAGNOSTIC;
        } else if (isOpaqueNestedRowMarkup(block.xml)) {
          markdown = "[Unsupported nested w:tr content]";
        } else {
          return { type: "delimiter" };
        }
        ctx.listIndentWidths = [];
        return { type: "paragraph", markdown };
      }
      // A delimiter, with no text to render.
      case "bookmarkStart":
      case "bookmarkEnd":
      case "moveFromRangeStart":
      case "moveFromRangeEnd":
      case "moveToRangeStart":
      case "moveToRangeEnd":
        return { type: "delimiter" };
      default: {
        const unsupported: never = block;
        panic(
          `Unsupported block content in markdown: ${JSON.stringify(unsupported)}`,
        );
      }
    }
  };

  // Every emitted block passes this boundary owner, including wrapper and
  // preserved-content projections. Only explicitly tight list items can join.
  for (const block of ordered) {
    const rendered = renderBlock(block);
    switch (rendered.type) {
      case "delimiter":
        continue;
      case "paragraph":
        if (rendered.markdown && out.length) out.push("");
        prevWasListItem = false;
        break;
      case "listItem":
        if (out.length && (!prevWasListItem || rendered.separation === "loose"))
          out.push("");
        prevWasListItem = true;
        break;
      default: {
        const unsupported: never = rendered;
        panic(`Unsupported rendered block: ${JSON.stringify(unsupported)}`);
      }
    }
    if (rendered.markdown) out.push(rendered.markdown);
  }

  ctx.listIndentWidths = outerListIndentWidths;
  return out.join("\n");
}
