/**
 * Page Break Extension — block node representing a DOCX page break
 */

import { createNodeExtension } from "../create";
import type { NodeExtension } from "../types";

/**
 * LEGACY_PAGE_BREAK_NODE_REMOVAL_CONDITION: delete the block node after persisted editor
 * snapshots created before `pageBreakRun` have been normalized and telemetry
 * confirms that no loaded snapshot contains a block-level page break. DOCX
 * import and editor commands must never create this legacy shape.
 */
export const PageBreakExtension: (
  options?: Partial<Record<string, unknown>> | undefined,
) => NodeExtension = createNodeExtension({
  name: "pageBreak",
  schemaNodeName: "pageBreak",
  nodeSpec: {
    group: "block",
    atom: true,
    selectable: true,
    parseDOM: [{ tag: "div.docx-page-break" }],
    toDOM() {
      return ["div", { class: "docx-page-break" }];
    },
  },
});
