/** Zero-width comment or tracked-move range that spans no content. */

import { readRangeAnchorAttrs, readRangeAnchorValues } from "../../rangeAnchorAttrs";
import { createNodeExtension } from "../create";

type RangeAnchorOptions = {
  getInternalClipboardToken?: () => string;
};

/** The node name, for callers asking whether a paragraph holds any content. */
export const RANGE_ANCHOR_NODE_NAME = "rangeAnchor";

/** The attribute the DOM spelling carries both markers in. */
export const RANGE_ANCHOR_DOM_ATTRIBUTE = "data-docx-range-anchor";

export const RangeAnchorExtension = createNodeExtension<RangeAnchorOptions>({
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
      start: {},
      end: {},
    },
    parseDOM: [
      {
        tag: `span[${RANGE_ANCHOR_DOM_ATTRIBUTE}]`,
        getAttrs(dom) {
          const raw = dom.getAttribute(RANGE_ANCHOR_DOM_ATTRIBUTE);
          if (raw === null) {
            return false;
          }
          const parsed = parseMarkers(raw);
          if (parsed === null) {
            return false;
          }
          const result = readRangeAnchorValues(parsed.start, parsed.end);
          return result.ok ? result.value : false;
        },
      },
    ],
    toDOM(node) {
      const attrs = readRangeAnchorAttrs(node);
      return [
        "span",
        {
          [RANGE_ANCHOR_DOM_ATTRIBUTE]: attrs.ok ? JSON.stringify(attrs.value) : "",
          "aria-hidden": "true",
          contenteditable: "false",
          style: "display: none;",
          ...(options.getInternalClipboardToken
            ? { "data-docx-internal-clipboard": options.getInternalClipboardToken() }
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
const parseMarkers = (raw: string): { start: unknown; end: unknown } | null => {
  try {
    const parsed: unknown = JSON.parse(raw);
    if (typeof parsed !== "object" || parsed === null) {
      return null;
    }
    if (!("start" in parsed) || !("end" in parsed)) {
      return null;
    }
    return { start: parsed.start, end: parsed.end };
  } catch {
    return null;
  }
};
