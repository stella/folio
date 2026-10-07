import { panic } from "better-result";
import { getCachedNumberingMap } from "../docx/numberingParser";
import { proseDocToBlocks } from "../prosemirror/conversion/fromProseDoc";
import { headerFooterToProseDoc } from "../prosemirror/conversion/toProseDoc";
import { createStyleEngine } from "../style-engine/styleEngine";
import type { BlockContent, DocxPackage } from "../types/document";
import { resolveWholeStory } from "./wholeStoryRevisionResolution";

/** Inline revisions alone are already handled by the Markdown run renderer. */
const hasBlockRevision = (blocks: readonly BlockContent[]): boolean =>
  blocks.some((block) => {
    switch (block.type) {
      case "paragraph":
        return block.pPrMark !== undefined;
      case "table":
        return block.rows.some(
          (row) =>
            row.structuralChange !== undefined ||
            row.cells.some(
              (cell) => cell.structuralChange !== undefined || hasBlockRevision(cell.content),
            ),
        );
      case "blockSdt":
      case "blockCustomXml":
        return hasBlockRevision(block.content);
      case "preservedBlock":
      case "bookmarkStart":
      case "bookmarkEnd":
        return false;
      default: {
        const unsupported: never = block;
        return panic("Unsupported block in accepted projection", { unsupported });
      }
    }
  });

/**
 * Read a model story through the same structural resolver as Accept All and
 * the editor's No Markup view. The authored model is never changed.
 */
export const projectAcceptedBlocks = (
  blocks: readonly BlockContent[],
  pkg: DocxPackage | undefined,
) => {
  if (!hasBlockRevision(blocks)) return null;
  const options = {
    ...(pkg?.styles !== undefined && { styles: pkg.styles }),
    ...(pkg?.theme !== undefined && { theme: pkg.theme }),
    ...(pkg?.numbering !== undefined && { numbering: pkg.numbering }),
  };
  const doc = headerFooterToProseDoc([...blocks], options);
  const result = resolveWholeStory({
    doc,
    mode: "accept",
    styleResolver: createStyleEngine(pkg?.styles),
    numbering: pkg?.numbering ? getCachedNumberingMap(pkg.numbering) : null,
  });
  return {
    blocks: result.resolved.eq(doc)
      ? blocks
      : proseDocToBlocks(result.resolved, [...blocks], pkg?.styles),
    completeness: result.failed ? "partial" : "complete",
  };
};
