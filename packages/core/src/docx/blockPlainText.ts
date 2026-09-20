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

import type { BlockContent } from "../types/document";
import { getParagraphText } from "./paragraphParser";

/** One entry per block, so callers can join or count lines as they need. */
export const collectBlockTexts = (blocks: readonly BlockContent[]): string[] => {
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
            row.cells.map((cell) => collectBlockTexts(cell.content).join("\n")).join("\t"),
          );
        }
        break;
      case "blockSdt":
        texts.push(...collectBlockTexts(block.content));
        break;
      // Opaque markup, so folio cannot say what text it puts on the page; an
      // entry of its own would claim an empty line the reader does not see.
      case "preservedBlock":
      // A delimiter, not content: it puts no text on the page.
      case "bookmarkStart":
      case "bookmarkEnd":
        break;
      default: {
        const unsupported: never = block;
        panic(`Unsupported block in plain-text extraction: ${JSON.stringify(unsupported)}`);
      }
    }
  }
  return texts;
};

export const blockPlainText = (blocks: readonly BlockContent[]): string =>
  collectBlockTexts(blocks).join("\n");
