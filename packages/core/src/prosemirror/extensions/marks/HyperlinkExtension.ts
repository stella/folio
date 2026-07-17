/**
 * Hyperlink Mark Extension
 */

import { panic } from "better-result";
import type { Command, EditorState } from "prosemirror-state";

import { expectHyperlinkMarkAttrs } from "../../attrs";
import { normalizeUserUrl, sanitizeExternalUrl } from "../../../utils/urlSecurity";
import { createMarkExtension } from "../create";
import type { ExtensionContext, ExtensionRuntime } from "../types";
import { isMarkActive } from "./markUtils";

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
      _docxHyperlinkIndex: { default: null },
    },
    inclusive: false,
    parseDOM: [
      {
        tag: "a[href]",
        getAttrs: (dom) => ({
          // HTMLElement.getAttribute is available on all element types.
          // Sanitize on the way in so pasted/programmatic anchors carrying
          // javascript:/data:/file: hrefs never make it into the mark.
          href: sanitizeStoredHref(dom.getAttribute("href") ?? undefined),
          tooltip: dom.getAttribute("title") ?? undefined,
        }),
      },
    ],
    toDOM(mark) {
      const { href, tooltip } = expectHyperlinkMarkAttrs(mark);
      const domAttrs: Record<string, string> = {
        // Defense in depth: re-sanitize the stored href before it reaches the
        // live DOM, in case a mark was created by another path.
        href: sanitizeStoredHref(href),
        target: "_blank",
        rel: "noopener noreferrer",
      };
      if (tooltip) {
        domAttrs["title"] = tooltip;
      }
      return ["a", domAttrs, 0];
    },
  },
  onSchemaReady(ctx: ExtensionContext): ExtensionRuntime {
    const hlType = ctx.schema.marks["hyperlink"];
    if (!hlType) {
      panic("Missing mark type: hyperlink");
    }

    const setHyperlink =
      (href: string, tooltip?: string): Command =>
      (state, dispatch) => {
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
      };

    const removeHyperlink: Command = (state, dispatch) => {
      const { from, to, empty } = state.selection;

      if (empty) {
        const $pos = state.selection.$from;
        const marks = $pos.marks();
        const linkMark = marks.find((m) => m.type === hlType);

        if (!linkMark) {
          return false;
        }

        let start = $pos.pos;
        let end = $pos.pos;

        const parent = $pos.parent;
        // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
        parent.forEach((node, offset) => {
          if (node.isText) {
            const nodeStart = $pos.start() + offset;
            const nodeEnd = nodeStart + node.nodeSize;

            if (nodeStart <= $pos.pos && $pos.pos <= nodeEnd) {
              const hasLink = node.marks.some((m) => m.type === hlType);
              if (hasLink) {
                start = Math.min(start, nodeStart);
                end = Math.max(end, nodeEnd);
              }
            }
          }
        });

        if (dispatch) {
          dispatch(state.tr.removeMark(start, end, hlType).scrollIntoView());
        }
        return true;
      }

      if (dispatch) {
        dispatch(state.tr.removeMark(from, to, hlType).scrollIntoView());
      }

      return true;
    };

    const insertHyperlink =
      (text: string, href: string, tooltip?: string): Command =>
      (state, dispatch) => {
        if (dispatch) {
          const mark = hlType.create({
            href: normalizeHyperlinkInput(href),
            tooltip: tooltip || null,
          });
          const textNode = state.schema.text(text, [mark]);
          dispatch(state.tr.replaceSelectionWith(textNode, false).scrollIntoView());
        }
        return true;
      };

    return {
      commands: {
        setHyperlink,
        removeHyperlink: () => removeHyperlink,
        insertHyperlink,
      },
    };
  },
});
