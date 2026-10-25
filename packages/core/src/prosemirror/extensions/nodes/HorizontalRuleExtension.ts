/**
 * Horizontal Rule Extension
 */

import { createNodeExtension } from "../create";
import type { NodeExtension } from "../types";

export const HorizontalRuleExtension: (
  options?: Partial<Record<string, unknown>> | undefined,
) => NodeExtension = createNodeExtension({
  name: "horizontalRule",
  schemaNodeName: "horizontalRule",
  nodeSpec: {
    group: "block",
    parseDOM: [{ tag: "hr" }],
    toDOM() {
      return ["hr"];
    },
  },
});
