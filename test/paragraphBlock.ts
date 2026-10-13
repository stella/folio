import { panic } from "better-result";
import type { FolioAIBlock } from "../packages/core/src/ai-edits/types";

/** A fixture expecting paragraph content must reject an opaque diagnostic carrier. */
export const expectParagraphBlock = (block: FolioAIBlock | undefined) => {
  if (block === undefined) return panic("The fixture has no paragraph block.");
  switch (block.kind) {
    case "paragraph":
    case "heading":
    case "listItem":
      return block;
    case "diagnostic":
      return panic("The fixture returned an opaque diagnostic instead of paragraph content.");
    default: {
      const unreachable: never = block;
      return panic("Unhandled fixture block kind", { block: unreachable });
    }
  }
};
