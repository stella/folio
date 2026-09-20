/**
 * The one place an inline-wrapper layer is built, and the codec that carries a
 * stack through the DOM.
 *
 * Two layers describing the same wrapper have to be the same value, because
 * `fromProseDoc` groups inline leaves into runs by `JSON.stringify` over their
 * mark attrs: one wrapper spelled with its keys in two orders would split one
 * run into two. Object key order is insertion order in JavaScript, so the
 * factory fixes it and every construction site goes through the factory —
 * including the one that reads a stack back out of the DOM, where `JSON.parse`
 * preserves whatever order the markup happened to carry.
 */

import { Result, panic } from "better-result";

import { readInlineWrapperStack } from "./attrs";
import type { InlineWrapperLayer } from "./schema/marks";

/** The DOM attribute that carries a serialised stack through copy and paste. */
export const INLINE_WRAPPER_STACK_ATTRIBUTE = "data-inline-wrapper";

/** A layer with its keys in canonical order, absent fields left absent. */
export const inlineWrapperLayer = (layer: InlineWrapperLayer): InlineWrapperLayer => {
  switch (layer.kind) {
    case "bidi":
      return layer.direction === undefined
        ? { kind: "bidi", control: layer.control }
        : { kind: "bidi", control: layer.control, direction: layer.direction };
    default: {
      // The union has one member today, so the value itself does not narrow to
      // `never`; the discriminant does, and it is what a new kind changes.
      layer.kind satisfies never;
      panic(`Unsupported inline wrapper layer: ${JSON.stringify(layer)}`);
    }
  }
};

/** `layers` in canonical form, or `null` when there is no wrapper to record. */
export const inlineWrapperStack = (
  layers: readonly InlineWrapperLayer[],
): readonly InlineWrapperLayer[] | null =>
  layers.length === 0 ? null : layers.map(inlineWrapperLayer);

export const serializeInlineWrapperStack = (stack: readonly InlineWrapperLayer[]): string =>
  JSON.stringify(stack);

/**
 * A stack read back from the DOM, or `null` when the markup carries none.
 *
 * Paste from outside the editor is the ordinary case: the attribute is absent,
 * so the pasted leaf carries no wrapper. Markup that states the attribute but
 * not a stack this build validates is read the same way — the text is what
 * matters, and inventing a wrapper around it would change the document.
 */
export const parseInlineWrapperStack = (
  serialized: string | null | undefined,
): readonly InlineWrapperLayer[] | null => {
  if (!serialized) {
    return null;
  }
  const parsed = Result.try((): unknown => JSON.parse(serialized));
  if (parsed.isErr()) {
    return null;
  }
  const stack = readInlineWrapperStack(parsed.value);
  return stack.ok ? stack.value.map(inlineWrapperLayer) : null;
};
