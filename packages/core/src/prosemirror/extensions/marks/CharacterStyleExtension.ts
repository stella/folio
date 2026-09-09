/**
 * Character Style Mark Extension (w:rStyle)
 *
 * Carries a run's character style reference through ProseMirror so semantic
 * styles (e.g. a house "DefinedTerm" style) survive the load → edit → save
 * round-trip. The style's resolved formatting is flattened into the regular
 * marks at load for rendering; this mark preserves only the reference. Style
 * formatting resolves through the document's shared style engine so it does
 * not become stale when a run moves between paragraph contexts.
 *
 * The mark renders as an unstyled span with a `data-character-style`
 * attribute. parseDOM restores the styleId (so copy/paste inside the editor
 * keeps the reference).
 */

import { expectCharacterStyleMarkAttrs } from "../../attrs";
import { createMarkExtension } from "../create";

export const CharacterStyleExtension = createMarkExtension({
  name: "characterStyle",
  schemaMarkName: "characterStyle",
  markSpec: {
    attrs: {
      styleId: { default: "" },
    },
    parseDOM: [
      {
        tag: "span[data-character-style]",
        getAttrs: (dom) => {
          // eslint-disable-next-line unicorn/prefer-dom-node-dataset -- getAttribute also works in XML/jsdom contexts where `dataset` may be absent.
          const styleId = dom.getAttribute("data-character-style");
          return styleId ? { styleId } : false;
        },
      },
    ],
    toDOM(mark) {
      const { styleId } = expectCharacterStyleMarkAttrs(mark);
      return ["span", { "data-character-style": styleId }, 0];
    },
  },
});
