/**
 * Doc Extension — top-level document node
 */

import { PROSE_PARAGRAPH_SOURCE_CONTRACT_ATTR } from "../../../docx/paragraphPropertySource";
import { createNodeExtension } from "../create";
import type { NodeExtension } from "../types";

export const DocExtension: (
  options?: Partial<Record<string, unknown>> | undefined,
) => NodeExtension = createNodeExtension({
  name: "doc",
  schemaNodeName: "doc",
  nodeSpec: {
    attrs: {
      [PROSE_PARAGRAPH_SOURCE_CONTRACT_ATTR]: { default: null },
      _finalSectionStart: { default: null },
      _adjustLineHeightInTable: { default: false },
      _doNotUseIndentAsNumberingTabStop: { default: false },
    },
    content:
      "(paragraph | horizontalRule | pageBreak | table | textBox | blockSdt | blockCustomXml | preservedBlock | blockBookmarkBoundary | blockMoveRangeBoundary)+",
  },
});
