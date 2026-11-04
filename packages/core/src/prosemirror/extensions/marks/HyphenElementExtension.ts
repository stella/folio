/**
 * Marks the native text that stands for a `w:softHyphen` or `w:noBreakHyphen`
 * run child (ECMA-376 Part 1, 17.3.3.29 and 17.3.3.18).
 *
 * The editor renders both elements as their Unicode characters so line
 * breaking and measurement treat them as text. The same characters may also be
 * authored inside `w:t`, so without this mark the save leg could not tell
 * which of the two the source held. Nothing is painted for it.
 *
 * Not inclusive, and each carrier in a paragraph has its own `id`: ProseMirror
 * keeps a non-inclusive mark only where both neighbours hold an equal mark, so
 * distinct ids mean a character typed beside or between carriers is text. The
 * ids only tell neighbours apart. They can repeat after a paste beside the
 * source or a paragraph join, which `hyphenElementRunContent` describes.
 *
 * External HTML must not be able to claim the element: `cleanPastedHtml` strips
 * the attribute unless the slice carries the internal clipboard capability, so
 * `toDOM` stamps that capability on every carrier it writes.
 */

import { expectHyphenElementMarkAttrs } from "../../attrs";
import { HYPHEN_ELEMENT_MARK_NAME } from "../../conversion/hyphenTextCarriers";
import { createMarkExtension } from "../create";
import type { MarkExtension } from "../types";

const HYPHEN_ELEMENT_ATTRIBUTE = "data-docx-hyphen-element";
const CANONICAL_HYPHEN_ELEMENT_ID = /^(?:0|[1-9]\d*)$/u;

type HyphenElementOptions = {
  getInternalClipboardToken?: () => string;
};

export const HyphenElementExtension: (
  options?: Partial<HyphenElementOptions> | undefined,
) => MarkExtension = createMarkExtension<HyphenElementOptions>({
  name: HYPHEN_ELEMENT_MARK_NAME,
  schemaMarkName: HYPHEN_ELEMENT_MARK_NAME,
  markSpec: (options) => ({
    attrs: { id: {} },
    inclusive: false,
    parseDOM: [
      {
        tag: `span[${HYPHEN_ELEMENT_ATTRIBUTE}]`,
        getAttrs(dom) {
          const rawId = dom.getAttribute(HYPHEN_ELEMENT_ATTRIBUTE);
          if (rawId === null || !CANONICAL_HYPHEN_ELEMENT_ID.test(rawId)) {
            return false;
          }
          const id = Number(rawId);
          return Number.isSafeInteger(id) ? { id } : false;
        },
      },
    ],
    toDOM(mark) {
      const { id } = expectHyphenElementMarkAttrs(mark);
      return [
        "span",
        {
          [HYPHEN_ELEMENT_ATTRIBUTE]: String(id),
          ...(options.getInternalClipboardToken
            ? {
                "data-docx-internal-clipboard":
                  options.getInternalClipboardToken(),
              }
            : {}),
        },
        0,
      ];
    },
  }),
});
