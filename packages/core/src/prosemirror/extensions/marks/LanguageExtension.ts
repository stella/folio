import { expectLanguageMarkAttrs } from "../../attrs";
import { createMarkExtension } from "../create";
import type { MarkExtension } from "../types";

/** Preserve OOXML `w:lang` and expose its primary language to the editing DOM. */
export const LanguageExtension: (
  options?: Partial<Record<string, unknown>> | undefined,
) => MarkExtension = createMarkExtension({
  name: "language",
  schemaMarkName: "language",
  markSpec: {
    attrs: {
      val: { default: null },
      eastAsia: { default: null },
      bidi: { default: null },
    },
    parseDOM: [
      {
        tag: "span[lang]",
        getAttrs: (dom) =>
          typeof dom === "string"
            ? false
            : { val: dom.getAttribute("lang") ?? undefined },
      },
    ],
    toDOM(mark) {
      const { val, eastAsia, bidi } = expectLanguageMarkAttrs(mark);
      const language = val ?? eastAsia ?? bidi;
      return language ? ["span", { lang: language }, 0] : ["span", 0];
    },
  },
});
