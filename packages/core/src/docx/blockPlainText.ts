/**
 * The one plain-text projection of a block sequence.
 *
 * Every story — body, header, footer, footnote, endnote — reads through here,
 * so a reader, a hash and a comparison all see the same string for the same
 * content. Header and footer text used to come from a separate walk that
 * handled only text runs, silently dropping fields, hyperlinks, tabs, breaks
 * and tracked changes, so the same paragraph read one way in a note and another
 * in a header.
 *
 * Separators carry the structure a reader needs: paragraphs and table rows are
 * newline-separated, cells within a row tab-separated. `getParagraphText` owns
 * everything below block level and reads the accepted tracked-change view.
 */

import { panic } from "better-result";

import { resolvedDocumentView } from "@stll/docx-core/ops";

import type { BlockContent, DocxPackage } from "../types/document";
import { ALT_CHUNK_READER_DIAGNOSTIC, isAltChunkMarkup } from "./altChunk";
import {
  isOpaqueNestedRowMarkup,
  OPAQUE_REVISION_CARRIER_READER_DIAGNOSTIC,
  opaqueRevisionCarrierName,
} from "./opaqueCarrier";
import { getParagraphText } from "./paragraphParser";

/** One entry per block, so callers can join or count lines as they need. */
const collectResolvedBlockTexts = (
  blocks: readonly BlockContent[],
): string[] => {
  const texts: string[] = [];
  for (const block of blocks) {
    switch (block.type) {
      case "paragraph":
        texts.push(getParagraphText(block));
        break;
      case "table":
        for (const row of block.rows) {
          // A hidden row is not part of what a reader sees.
          if (row.formatting?.hidden === true) {
            continue;
          }
          texts.push(
            row.cells
              .map((cell) => collectResolvedBlockTexts(cell.content).join("\n"))
              .join("\t"),
          );
        }
        break;
      case "blockSdt":
      case "blockCustomXml":
        texts.push(...collectResolvedBlockTexts(block.content));
        break;
      case "preservedBlock":
        if (isAltChunkMarkup(block.xml)) {
          texts.push(
            block.readerText === undefined
              ? ALT_CHUNK_READER_DIAGNOSTIC
              : `${ALT_CHUNK_READER_DIAGNOSTIC}\n${block.readerText}`,
          );
        } else if (opaqueRevisionCarrierName(block.xml) !== undefined) {
          texts.push(OPAQUE_REVISION_CARRIER_READER_DIAGNOSTIC);
        } else if (isOpaqueNestedRowMarkup(block.xml)) {
          texts.push("[Unsupported nested w:tr content]");
        }
        break;
      // Opaque markup, so folio cannot say what text it puts on the page; an
      // entry of its own would claim an empty line the reader does not see.
      // A delimiter, not content: it puts no text on the page.
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
        const unsupported: never = block;
        panic(
          `Unsupported block in plain-text extraction: ${JSON.stringify(unsupported)}`,
        );
      }
    }
  }
  return texts;
};

/** The kind of note whose automatic mark the blocks hold, if any. */
const noteMarkerKind = (value: unknown): "footnote" | "endnote" | undefined => {
  if (Array.isArray(value)) {
    for (const item of value) {
      const kind = noteMarkerKind(item);
      if (kind !== undefined) return kind;
    }
    return undefined;
  }
  if (typeof value !== "object" || value === null) return undefined;
  if ("type" in value && value.type === "noteMarker" && "kind" in value)
    return value.kind === "endnote" ? "endnote" : "footnote";
  for (const field of ["content", "rows", "cells"] as const) {
    if (field in value) {
      const kind = noteMarkerKind((value as Record<string, unknown>)[field]);
      if (kind !== undefined) return kind;
    }
  }
  return undefined;
};

/** The blocks accepted as Accept All accepts them, read as the story they belong to. */
const acceptedBlocks = (blocks: readonly BlockContent[]) => {
  // A note's blocks hold its automatic mark, which only a note story of
  // that kind may: they resolve as that note, not as a body.
  const kind = noteMarkerKind(blocks);
  if (kind === undefined) {
    const view = resolvedDocumentView(
      { package: { document: { content: [...blocks] } } },
      "accept",
    );
    return { blocks: view.document.package.document.content, view };
  }
  const content: BlockContent[] = [...blocks];
  const notePackage: DocxPackage =
    kind === "footnote"
      ? {
          document: { content: [] },
          footnotes: [{ type: "footnote", id: 1, content }],
        }
      : {
          document: { content: [] },
          endnotes: [{ type: "endnote", id: 1, content }],
        };
  const view = resolvedDocumentView({ package: notePackage }, "accept");
  const resolved =
    kind === "footnote"
      ? view.document.package.footnotes?.[0]
      : view.document.package.endnotes?.[0];
  return { blocks: resolved?.content ?? [], view };
};

/** One entry per accepted block; deleted breaks and table structure resolve together. */
export const collectBlockTexts = (
  blocks: readonly BlockContent[],
): string[] => {
  // The blocks read as one story, accepted as Accept All accepts it.
  const { blocks: accepted, view } = acceptedBlocks(blocks);
  const texts = collectResolvedBlockTexts(accepted);
  if (view.unresolved.length > 0) {
    texts.push("[Unresolved tracked structural changes]");
  }
  return texts;
};

export const blockPlainText = (blocks: readonly BlockContent[]): string =>
  collectBlockTexts(blocks).join("\n");
