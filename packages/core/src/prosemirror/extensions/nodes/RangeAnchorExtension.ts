/** Zero-width comment or tracked-move range that spans no content. */

import { MAX_REVIEW_HISTORY_ATTRIBUTE_LENGTH } from "../../../reviewHistoryLimits";
import {
  INTERNAL_CLIPBOARD_TOKEN_ATTRIBUTE,
  isInternalClipboardDom,
} from "../../copyOriginDom";
import { withoutUntrustedCopyOrigin } from "../../moveRangeBoundaryAttrs";
import {
  readRangeAnchorAttrs,
  readRangeAnchorValues,
} from "../../rangeAnchorAttrs";
import { createNodeExtension } from "../create";
import type { NodeExtension } from "../types";
import { INLINE_ANCHOR_DOM_STYLE } from "./inlineAnchorDom";

type RangeAnchorOptions = {
  getInternalClipboardToken?: () => string;
};

/** The node name, for callers asking whether a paragraph holds any content. */
export const RANGE_ANCHOR_NODE_NAME = "rangeAnchor";

/** The attribute the DOM spelling carries both markers in. */
export const RANGE_ANCHOR_DOM_ATTRIBUTE = "data-docx-range-anchor";

export const RangeAnchorExtension: (
  options?: Partial<RangeAnchorOptions> | undefined,
) => NodeExtension = createNodeExtension<RangeAnchorOptions>({
  name: RANGE_ANCHOR_NODE_NAME,
  schemaNodeName: RANGE_ANCHOR_NODE_NAME,
  nodeSpec: (options) => ({
    inline: true,
    group: "inline",
    marks: "_",
    atom: true,
    // A point comment stays a point unless the reviewer selects text: there is
    // nothing here to select, and nothing a caret can sit inside, so typing
    // beside the anchor leaves the range exactly as wide as it was.
    selectable: false,
    attrs: {
      markers: {},
    },
    parseDOM: [
      {
        tag: `span[${RANGE_ANCHOR_DOM_ATTRIBUTE}]`,
        getAttrs(dom) {
          const raw = dom.getAttribute(RANGE_ANCHOR_DOM_ATTRIBUTE);
          if (
            raw === null ||
            raw.length > MAX_REVIEW_HISTORY_ATTRIBUTE_LENGTH
          ) {
            return false;
          }
          const parsed = parseMarkers(raw);
          if (parsed === null) {
            return false;
          }
          const markers = isInternalClipboardDom(
            dom,
            options.getInternalClipboardToken?.(),
          )
            ? parsed
            : { markers: parsed.markers.map(withoutUntrustedCopyOrigin) };
          const result = readRangeAnchorValues(markers.markers);
          return result.ok ? result.value : false;
        },
      },
    ],
    toDOM(node) {
      const attrs = readRangeAnchorAttrs(node);
      const internalClipboardToken = options.getInternalClipboardToken?.();
      let serializedMarkers = "";
      if (attrs.ok) {
        const markers =
          internalClipboardToken === undefined
            ? { markers: attrs.value.markers.map(withoutUntrustedCopyOrigin) }
            : attrs.value;
        serializedMarkers = JSON.stringify(markers);
      }
      return [
        "span",
        {
          [RANGE_ANCHOR_DOM_ATTRIBUTE]: serializedMarkers,
          "aria-hidden": "true",
          contenteditable: "false",
          style: INLINE_ANCHOR_DOM_STYLE,
          ...(internalClipboardToken
            ? {
                [INTERNAL_CLIPBOARD_TOKEN_ATTRIBUTE]: internalClipboardToken,
              }
            : {}),
        },
      ];
    },
  }),
});

/**
 * The attribute back as an attrs bag, or `null` when it is not one.
 *
 * `JSON.parse` throws on malformed input reaching the schema from a paste, and
 * a parser branch that refuses the node is the boundary this belongs at.
 */
const parseMarkers = (raw: string) => {
  try {
    const parsed: unknown = JSON.parse(raw);
    if (typeof parsed !== "object" || parsed === null || !("markers" in parsed))
      return null;
    const result = readRangeAnchorValues(parsed.markers);
    return result.ok ? result.value : null;
  } catch {
    return null;
  }
};
