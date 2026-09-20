/**
 * Inline wrapper mark — the transparent wrappers a leaf sits inside.
 *
 * One mark for every wrapper kind rather than one mark per kind: ProseMirror's
 * mark set is unordered across types, so two marks could not say which wrapper
 * is inside which. The nesting lives in the `stack` attr, outermost first, and
 * the mark excludes itself (ProseMirror's default) because a second one would
 * be a second answer to the same question.
 *
 * `inclusive: false` so typing at the end of a wrapped span does not silently
 * extend the wrapper over text the author wrote outside it.
 */

import { panic } from "better-result";
import type { DOMOutputSpec } from "prosemirror-model";

import { expectInlineWrapperMarkAttrs } from "../../attrs";
import {
  INLINE_WRAPPER_STACK_ATTRIBUTE,
  parseInlineWrapperStack,
  serializeInlineWrapperStack,
} from "../../inlineWrapperStack";
import type { InlineWrapperLayer } from "../../schema/marks";
import { createMarkExtension } from "../create";

export const INLINE_WRAPPER_MARK_NAME = "inlineWrapper";

/** What a tagged layer's element name is spelled as on the rendered span. */
const TAGGED_ELEMENT_ATTRIBUTES = {
  smartTag: "data-smart-tag-element",
  customXml: "data-custom-xml-element",
} as const;

/**
 * How one layer is spelled in HTML.
 *
 * `w:bdo` is `unicode-bidi: bidi-override`, which HTML gives an element of its
 * own; `w:dir` is `unicode-bidi: embed`, which a `dir`-carrying span states.
 * A smart tag and a custom-XML wrapper say nothing about layout, so they are a
 * passthrough `<span>` that names the tagged element for a stylesheet or a
 * host that wants to show it; the stack attribute is what carries them back.
 */
const layerElement = (
  layer: InlineWrapperLayer,
  attributes: Record<string, string>,
): DOMOutputSpec => {
  switch (layer.kind) {
    case "bidi": {
      if (layer.direction !== undefined) {
        attributes["dir"] = layer.direction;
      }
      return [layer.control === "override" ? "bdo" : "span", attributes, 0];
    }
    case "smartTag":
    case "customXml":
      attributes[TAGGED_ELEMENT_ATTRIBUTES[layer.kind]] = layer.element;
      return ["span", attributes, 0];
    default: {
      layer satisfies never;
      panic(`Unspelled inline wrapper layer: ${JSON.stringify(layer)}`);
    }
  }
};

export const InlineWrapperExtension = createMarkExtension({
  name: INLINE_WRAPPER_MARK_NAME,
  schemaMarkName: INLINE_WRAPPER_MARK_NAME,
  markSpec: {
    attrs: {
      /** Outermost first. `null` is no wrapper, which is the absence of the mark. */
      stack: { default: null },
    },
    inclusive: false,
    parseDOM: [
      {
        tag: `[${INLINE_WRAPPER_STACK_ATTRIBUTE}]`,
        getAttrs(dom) {
          const stack = parseInlineWrapperStack(dom.getAttribute(INLINE_WRAPPER_STACK_ATTRIBUTE));
          return stack === null ? false : { stack };
        },
      },
    ],
    toDOM(mark) {
      const { stack } = expectInlineWrapperMarkAttrs(mark);
      // The DOM nests one element per mark, so only the innermost layer can be
      // spelled as an element; the attribute carries the whole stack back.
      const innermost = stack.at(-1);
      if (innermost === undefined) {
        panic("An inline wrapper mark carries at least one layer");
      }
      return layerElement(innermost, {
        [INLINE_WRAPPER_STACK_ATTRIBUTE]: serializeInlineWrapperStack(stack),
      });
    },
  },
});
