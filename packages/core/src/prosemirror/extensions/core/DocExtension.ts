/**
 * Doc Extension — top-level document node
 */

import { createNodeExtension } from "../create";
import { PROSE_PARAGRAPH_SOURCE_CONTRACT_ATTR } from "../../../docx/paragraphPropertySource";

export const DocExtension = createNodeExtension({
  name: "doc",
  schemaNodeName: "doc",
  nodeSpec: {
    attrs: {
      [PROSE_PARAGRAPH_SOURCE_CONTRACT_ATTR]: { default: null },
      _finalSectionStart: { default: null },
      _adjustLineHeightInTable: { default: false },
    },
    content: "(paragraph | horizontalRule | pageBreak | table | textBox | blockSdt)+",
  },
});
