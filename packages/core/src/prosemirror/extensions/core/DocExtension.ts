/**
 * Doc Extension — top-level document node
 */

import { createNodeExtension } from "../create";

export const DocExtension = createNodeExtension({
  name: "doc",
  schemaNodeName: "doc",
  nodeSpec: {
    attrs: {
      _finalSectionStart: { default: null },
    },
    content: "(paragraph | horizontalRule | pageBreak | table | textBox | blockSdt)+",
  },
});
