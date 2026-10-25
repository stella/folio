/**
 * Text Extension — inline text node
 */

import { createNodeExtension } from "../create";
import type { NodeExtension } from "../types";

export const TextExtension: (
  options?: Partial<Record<string, unknown>> | undefined,
) => NodeExtension = createNodeExtension({
  name: "text",
  schemaNodeName: "text",
  nodeSpec: {
    group: "inline",
  },
});
