/**
 * All Caps Mark Extension (w:caps)
 */

import { createMarkExtension } from "../create";
import type { MarkExtension } from "../types";

export const AllCapsExtension: (
  options?: Partial<Record<string, unknown>> | undefined,
) => MarkExtension = createMarkExtension({
  name: "allCaps",
  schemaMarkName: "allCaps",
  markSpec: {
    parseDOM: [
      {
        style: "text-transform",
        getAttrs: (value) => (value === "uppercase" ? {} : false),
      },
    ],
    toDOM() {
      return ["span", { style: "text-transform: uppercase" }, 0];
    },
  },
});
