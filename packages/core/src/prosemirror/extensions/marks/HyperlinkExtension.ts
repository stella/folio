/**
 * Hyperlink Mark Extension
 */

import { withCanonicalCommand } from "../../canonicalCommands";
import { BUILT_IN_STYLE_NAME } from "../../../docx/builtInStyles";
import { getDocumentBuiltInStyles } from "../../plugins/documentStyles";
import { panic } from "better-result";
import type { Command, EditorState } from "prosemirror-state";

import { expectHyperlinkMarkAttrs } from "../../attrs";
import {
  anchorTargetAttrs,
  normalizeUserUrl,
  sanitizeExternalUrl,
} from "../../../utils/urlSecurity";
import { removeHyperlinkInRange } from "../../hyperlinkRemoval";
import { createMarkExtension } from "../create";
import type { ExtensionContext, ExtensionRuntime } from "../types";
import { isMarkActive } from "./markUtils";

const DOCX_HYPERLINK_INDEX_ATTRIBUTE = "data-docx-hyperlink-index";

// ============================================================================
// HREF SANITIZATION HELPERS
// ============================================================================

// Internal bookmark anchors (`#name`) are not absolute URLs, so they never
// reach `sanitizeExternalUrl`/`normalizeUserUrl` (which parse via the `URL`
// constructor and would reject them). Adapters resolve `#name` hrefs to
// bookmark navigation directly off the DOM attribute, so they must survive
// both the DOM round-trip and dialog/programmatic writes unchanged.

/**
 * Sanitize an href arriving from (or being emitted to) the DOM: pasted HTML,
 * clipboard content, or a mark attr already stored on the document. Only
 * allow-listed absolute URLs (http/https/mailto/tel) and internal bookmark
 * anchors survive; everything else (javascript:, data:, file:, ...) becomes
 * an empty href, matching how the codebase already signals "no link" for an
 * unresolved hyperlink (see `toProseDoc.ts`'s `hyperlink.href || ""`).
 */
function sanitizeStoredHref(rawHref: string | undefined): string {
  if (!rawHref) {
    return "";
  }
  const trimmed = rawHref.trim();
  if (trimmed.startsWith("#")) {
    return trimmed;
  }
  return sanitizeExternalUrl(trimmed) ?? "";
}

/**
 * Normalize/sanitize a URL typed by a user (hyperlink dialog/popup) before it
 * is written into a mark. Protocol-less input (e.g. "example.com") is
 * accepted and normalized to https; anything resolving to a disallowed
 * scheme is dropped.
 */
function normalizeHyperlinkInput(rawHref: string): string {
  const trimmed = rawHref.trim();
  if (trimmed.startsWith("#")) {
    return trimmed;
  }
  return normalizeUserUrl(trimmed);
}

// ============================================================================
// HYPERLINK QUERY HELPERS (exported for toolbar)
// ============================================================================

export function isHyperlinkActive(state: EditorState): boolean {
  const hlType = state.schema.marks["hyperlink"];
  if (!hlType) {
    return false;
  }
  return isMarkActive(state, hlType);
}

export function getHyperlinkAttrs(state: EditorState): { href: string; tooltip?: string } | null {
  const hlType = state.schema.marks["hyperlink"];
  if (!hlType) {
    return null;
  }

  const { empty, $from, from, to } = state.selection;

  if (empty) {
    const marks = state.storedMarks ?? $from.marks();
    for (const mark of marks) {
      if (mark.type === hlType) {
        const { href, tooltip } = expectHyperlinkMarkAttrs(mark);
        return { href, ...(tooltip !== undefined ? { tooltip } : {}) };
      }
    }
    return null;
  }

  let attrs: { href: string; tooltip?: string } | null = null;
  state.doc.nodesBetween(from, to, (node) => {
    if (node.isText && attrs === null) {
      const mark = hlType.isInSet(node.marks);
      if (mark) {
        const { href, tooltip } = expectHyperlinkMarkAttrs(mark);
        attrs = { href, ...(tooltip !== undefined ? { tooltip } : {}) };
        return false;
      }
    }
    return true;
  });

  return attrs;
}

export function getSelectedText(state: EditorState): string {
  const { from, to, empty } = state.selection;
  if (empty) {
    return "";
  }
  return state.doc.textBetween(from, to, "");
}

// ============================================================================
// EXTENSION
// ============================================================================

export const HyperlinkExtension = createMarkExtension({
  name: "hyperlink",
  schemaMarkName: "hyperlink",
  markSpec: {
    attrs: {
      href: {},
      tooltip: { default: null },
      rId: { default: null },
      target: { default: null },
      history: { default: null },
      docLocation: { default: null },
      _docxHyperlinkIndex: { default: null },
    },
    inclusive: false,
    parseDOM: [
      {
        tag: "a[href]",
        getAttrs: (dom) => {
          const index = dom.getAttribute(DOCX_HYPERLINK_INDEX_ATTRIBUTE);
          const docxHyperlinkIndex =
            index !== null && /^\d+$/u.test(index) && Number.isSafeInteger(Number(index))
              ? Number(index)
              : undefined;
          return {
            // HTMLElement.getAttribute is available on all element types.
            // Sanitize on the way in so pasted/programmatic anchors carrying
            // javascript:/data:/file: hrefs never make it into the mark.
            href: sanitizeStoredHref(dom.getAttribute("href") ?? undefined),
            tooltip: dom.getAttribute("title") ?? undefined,
            target: dom.getAttribute("target") ?? undefined,
            ...(docxHyperlinkIndex === undefined
              ? {}
              : { _docxHyperlinkIndex: docxHyperlinkIndex }),
          };
        },
      },
    ],
    toDOM(mark) {
      const { href, target, tooltip, _docxHyperlinkIndex } = expectHyperlinkMarkAttrs(mark);
      const domAttrs: Record<string, string> = {
        // Defense in depth: re-sanitize the stored href before it reaches the
        // live DOM, in case a mark was created by another path.
        href: sanitizeStoredHref(href),
        ...anchorTargetAttrs(target),
      };
      if (tooltip) {
        domAttrs["title"] = tooltip;
      }
      if (typeof _docxHyperlinkIndex === "number") {
        domAttrs[DOCX_HYPERLINK_INDEX_ATTRIBUTE] = String(_docxHyperlinkIndex);
      }
      return ["a", domAttrs, 0];
    },
  },
  onSchemaReady(ctx: ExtensionContext): ExtensionRuntime {
    if (!ctx.schema.marks["hyperlink"]) {
      panic("Missing mark type: hyperlink");
    }
    // Resolved per call: the document may be of another schema instance.
    const documentHyperlinkType = (state: EditorState) =>
      state.schema.marks["hyperlink"] ?? panic("Missing mark type: hyperlink");

    const setHyperlink = (href: string, tooltip?: string): Command =>
      withCanonicalCommand(
        (state, dispatch) => {
          const hlType = documentHyperlinkType(state);
          const { from, to, empty } = state.selection;

          if (empty) {
            return false;
          }

          if (dispatch) {
            const mark = hlType.create({
              href: normalizeHyperlinkInput(href),
              tooltip: tooltip || null,
            });
            let tr = state.tr.addMark(from, to, mark);
            // Remove any explicit text color so the default hyperlink blue (#0563c1)
            // shows through, matching MS Word behavior
            const textColorType = state.schema.marks["textColor"];
            if (textColorType) {
              tr = tr.removeMark(from, to, textColorType);
            }
            dispatch(tr.scrollIntoView());
          }

          return true;
        },
        (state) =>
          state.selection.empty
            ? []
            : [
                {
                  type: "setHyperlink",
                  from: state.selection.from,
                  to: state.selection.to,
                  href: normalizeHyperlinkInput(href),
                  ...(tooltip ? { tooltip } : {}),
                },
              ],
      );

    const removalRange = (state: EditorState) => {
      const { from, to, empty, $from } = state.selection;
      if (!empty) return { from, to };
      const hlType = documentHyperlinkType(state);
      if (!$from.marks().some((mark) => mark.type === hlType)) return undefined;
      let start = from;
      let end = to;
      $from.parent.forEach((node, offset) => {
        const nodeStart = $from.start() + offset;
        const nodeEnd = nodeStart + node.nodeSize;
        if (
          node.isText &&
          nodeStart <= from &&
          from <= nodeEnd &&
          node.marks.some((mark) => mark.type === hlType)
        ) {
          start = Math.min(start, nodeStart);
          end = Math.max(end, nodeEnd);
        }
      });
      return { from: start, to: end };
    };
    const canonicalRemovalRange = (state: EditorState) => {
      const { from, to, empty, $from } = state.selection;
      if (!empty) return { from, to };
      const linkMark = $from.marks().find((mark) => mark.type === documentHyperlinkType(state));
      if (!linkMark) return undefined;
      let contiguous: { from: number; to: number } | undefined;
      let selected: typeof contiguous;
      $from.parent.forEach((node, offset) => {
        if (!node.isText || !node.marks.some((mark) => mark.eq(linkMark))) {
          contiguous = undefined;
          return;
        }
        const start = $from.start() + offset;
        const end = start + node.nodeSize;
        if (contiguous?.to === start) contiguous.to = end;
        else contiguous = { from: start, to: end };
        if (start <= from && from <= end) selected = contiguous;
      });
      return selected;
    };
    const removeHyperlink = withCanonicalCommand(
      (state, dispatch) => {
        const range = removalRange(state);
        if (!range) return false;
        if (dispatch)
          dispatch(removeHyperlinkInRange(state, state.tr, range.from, range.to).scrollIntoView());
        return true;
      },
      (state) => {
        const range = canonicalRemovalRange(state);
        if (!range) return [];
        const hyperlinkStyleId = getDocumentBuiltInStyles(state).styleIdForBuiltInName(
          BUILT_IN_STYLE_NAME.hyperlink,
        );
        return [
          {
            type: "removeHyperlink",
            ...range,
            ...(hyperlinkStyleId === undefined ? {} : { hyperlinkStyleId }),
          },
        ];
      },
    );

    const insertHyperlink = (text: string, href: string, tooltip?: string): Command =>
      withCanonicalCommand(
        (state, dispatch) => {
          const hlType = documentHyperlinkType(state);
          if (dispatch) {
            const mark = hlType.create({
              href: normalizeHyperlinkInput(href),
              tooltip: tooltip || null,
            });
            const textNode = state.schema.text(text, [mark]);
            dispatch(state.tr.replaceSelectionWith(textNode, false).scrollIntoView());
          }
          return true;
        },
        (state) => [
          {
            type: "insertHyperlink",
            from: state.selection.from,
            to: state.selection.to,
            text,
            href: normalizeHyperlinkInput(href),
            ...(tooltip ? { tooltip } : {}),
          },
        ],
      );

    return {
      commands: {
        setHyperlink,
        removeHyperlink: () => removeHyperlink,
        insertHyperlink,
      },
    };
  },
});
