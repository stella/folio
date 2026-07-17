/**
 * Hyperlink popup commands
 *
 * Framework-agnostic ProseMirror operations behind the Google Docs-style
 * hyperlink popup (edit text + URL, remove). Extracted from the React
 * `useHyperlinkHandlers` hook so the adapter keeps only the popup state and the
 * browser glue (navigate, clipboard copy, toast).
 *
 * Both operate on the contiguous run of text nodes around the cursor that carry
 * a hyperlink mark with the same href, mirroring how Word/Docs treat a link as
 * one unit.
 */

import type { Mark, MarkType, Node as PMNode } from "prosemirror-model";
import type { EditorView } from "prosemirror-view";

import { expectHyperlinkMarkAttrs } from "../attrs";
import { normalizeUserUrl } from "../../utils/urlSecurity";

type HyperlinkRange = { start: number; end: number };

const hyperlinkHref = (mark: Mark): string => expectHyperlinkMarkAttrs(mark).href;

/**
 * Normalize/sanitize a URL typed by a user (hyperlink popup edit) before it
 * is written into a mark. Internal bookmark anchors (`#name`) bypass
 * `normalizeUserUrl` (which parses via the `URL` constructor and would
 * reject a bare fragment); protocol-less input (e.g. "example.com") is
 * accepted and normalized to https; anything resolving to a disallowed
 * scheme (javascript:, data:, file:, ...) is dropped to an empty href.
 */
const normalizeHyperlinkInput = (rawHref: string): string => {
  const trimmed = rawHref.trim();
  if (trimmed.startsWith("#")) {
    return trimmed;
  }
  return normalizeUserUrl(trimmed);
};

/**
 * Contiguous ranges of text nodes in `parent` (starting at `parentStart`) that
 * carry a hyperlink mark with the given `href`.
 */
const collectHyperlinkRanges = (
  parent: PMNode,
  parentStart: number,
  hlType: MarkType,
  href: string,
): HyperlinkRange[] => {
  const ranges: HyperlinkRange[] = [];
  let current: HyperlinkRange | null = null;

  // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node API
  parent.forEach((node, offset) => {
    const nodeStart = parentStart + offset;
    const nodeEnd = nodeStart + node.nodeSize;
    const hasMatchingLink = node.isText
      ? node.marks.some((m) => m.type === hlType && hyperlinkHref(m) === href)
      : false;

    if (hasMatchingLink) {
      if (current) {
        current.end = nodeEnd;
      } else {
        current = { start: nodeStart, end: nodeEnd };
      }
    } else if (current) {
      ranges.push(current);
      current = null;
    }
  });
  if (current) {
    ranges.push(current);
  }
  return ranges;
};

const rangeAtCursor = (ranges: HyperlinkRange[], cursorPos: number): HyperlinkRange | undefined =>
  ranges.find((range) => range.start <= cursorPos && cursorPos <= range.end);

export type EditHyperlinkInput = {
  displayText: string;
  href: string;
};

/**
 * Replace the text and href of the hyperlink at the cursor.
 *
 * Returns whether the adapter should close the popup and refocus the editor:
 * true after a successful edit and when the schema has no hyperlink behaviour
 * to act on at the cursor; false only when a hyperlink mark is present but no
 * contiguous range contains the cursor (the original's silent bail).
 */
export const editHyperlinkAtCursor = (
  view: EditorView,
  { displayText, href }: EditHyperlinkInput,
): boolean => {
  const hlType = view.state.schema.marks["hyperlink"];
  if (!hlType) {
    return false;
  }

  const { $from } = view.state.selection;
  const linkMark = $from.marks().find((m) => m.type === hlType);
  if (!linkMark) {
    return true;
  }

  const ranges = collectHyperlinkRanges(
    $from.parent,
    $from.start(),
    hlType,
    hyperlinkHref(linkMark),
  );
  const targetRange = rangeAtCursor(ranges, $from.pos);
  if (!targetRange) {
    return false;
  }

  const { tooltip } = expectHyperlinkMarkAttrs(linkMark);
  const newMark = hlType.create({ href: normalizeHyperlinkInput(href), tooltip });
  const textNode = view.state.schema.text(displayText, [
    ...$from.marks().filter((m) => m.type !== hlType),
    newMark,
  ]);
  const tr = view.state.tr.replaceWith(targetRange.start, targetRange.end, textNode);
  view.dispatch(tr.scrollIntoView());
  return true;
};

export type RemoveHyperlinkInput = {
  /** href from the popup, used as a fallback when the cursor reports no mark. */
  popupHref: string | undefined;
};

/**
 * Remove the hyperlink mark from the contiguous run at the cursor. Returns
 * whether a link was actually removed (so the adapter can toast + refocus).
 */
export const removeHyperlinkAtCursor = (
  view: EditorView,
  { popupHref }: RemoveHyperlinkInput,
): boolean => {
  const hlType = view.state.schema.marks["hyperlink"];
  if (!hlType) {
    return false;
  }

  const { $from } = view.state.selection;

  // Marks may not be reported exactly at a boundary position, so check the
  // cursor's own marks, then the nodes either side, then fall back to the
  // popup's href.
  let linkMark = $from.marks().find((m) => m.type === hlType);
  if (!linkMark && $from.nodeAfter) {
    linkMark = $from.nodeAfter.marks.find((m) => m.type === hlType);
  }
  if (!linkMark && $from.nodeBefore) {
    linkMark = $from.nodeBefore.marks.find((m) => m.type === hlType);
  }

  if (!linkMark && popupHref !== undefined) {
    // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node API
    $from.parent.forEach((node) => {
      if (!linkMark && node.isText) {
        const found = node.marks.find((m) => m.type === hlType && hyperlinkHref(m) === popupHref);
        if (found) {
          linkMark = found;
        }
      }
    });
  }

  if (!linkMark) {
    return false;
  }

  const ranges = collectHyperlinkRanges(
    $from.parent,
    $from.start(),
    hlType,
    hyperlinkHref(linkMark),
  );
  const targetRange = rangeAtCursor(ranges, $from.pos);
  if (!targetRange) {
    return false;
  }

  const tr = view.state.tr.removeMark(targetRange.start, targetRange.end, hlType);
  view.dispatch(tr.scrollIntoView());
  return true;
};
