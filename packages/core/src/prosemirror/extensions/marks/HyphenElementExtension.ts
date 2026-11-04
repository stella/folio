/**
 * Marks the native text that stands for a `w:softHyphen` or `w:noBreakHyphen`
 * run child (ECMA-376 Part 1, 17.3.3.29 and 17.3.3.18).
 *
 * The editor renders both elements as their Unicode characters so line
 * breaking and measurement treat them as text. The same characters may also be
 * authored inside `w:t`, so without this mark the save leg could not tell
 * which of the two the source held. Nothing is painted for it.
 *
 * Not inclusive: text typed beside a carrier is text, not another element.
 */

import { HYPHEN_ELEMENT_MARK_NAME } from "../../conversion/hyphenTextCarriers";
import { createMarkExtension } from "../create";
import type { MarkExtension } from "../types";

const HYPHEN_ELEMENT_ATTRIBUTE = "data-docx-hyphen-element";

export const HyphenElementExtension: (
  options?: Partial<Record<string, unknown>> | undefined,
) => MarkExtension = createMarkExtension({
  name: HYPHEN_ELEMENT_MARK_NAME,
  schemaMarkName: HYPHEN_ELEMENT_MARK_NAME,
  markSpec: {
    inclusive: false,
    parseDOM: [{ tag: `span[${HYPHEN_ELEMENT_ATTRIBUTE}]` }],
    toDOM() {
      return ["span", { [HYPHEN_ELEMENT_ATTRIBUTE]: "" }, 0];
    },
  },
});
