import type { Node as PMNode } from "prosemirror-model";

import { expectTextBoxAttrs } from "./attrs";

const hostParagraphs = new WeakMap<PMNode, PMNode | null>();

/**
 * The `w:p` a standalone text box was lifted out of, as a childless paragraph
 * node over the host's paragraph attributes, or `undefined` when the box
 * carries none.
 *
 * Building the node runs the carried attributes through the paragraph schema,
 * so every consumer reads them with the validation any paragraph's get.
 */
export const textBoxHostParagraph = (textBox: PMNode): PMNode | undefined => {
  const cached = hostParagraphs.get(textBox);
  if (cached !== undefined) {
    return cached ?? undefined;
  }
  const host = expectTextBoxAttrs(textBox)._docxHostParagraph;
  const paragraph = host ? textBox.type.schema.nodes["paragraph"]?.create(host) : undefined;
  hostParagraphs.set(textBox, paragraph ?? null);
  return paragraph;
};
