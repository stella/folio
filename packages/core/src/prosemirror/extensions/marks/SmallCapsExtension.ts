/**
 * Small Caps Mark Extension (w:smallCaps)
 */

import { createMarkExtension } from "../create";
import type { MarkExtension } from "../types";

export const SmallCapsExtension: (
  options?: Partial<Record<string, unknown>> | undefined,
) => MarkExtension = createMarkExtension({
  name: "smallCaps",
  schemaMarkName: "smallCaps",
  markSpec: {
    parseDOM: [
      {
        style: "font-variant",
        getAttrs: (value) => (value === "small-caps" ? {} : false),
      },
    ],
    toDOM() {
      return ["span", { style: "font-variant: small-caps" }, 0];
    },
  },
});
