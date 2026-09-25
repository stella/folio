/**
 * A paragraph's logical offset space.
 *
 * Every position an operation names is a count of units from the start of a
 * paragraph:
 *
 * - a character of run text is one unit per UTF-16 code unit;
 * - every other run child (tab, break, symbol, note reference, drawing,
 *   captured markup, field character) is one unit, except the cached
 *   rendered-page-break marker a layout pass writes, which is zero-width;
 * - a field, an equation and a comment reference are one unit each, however
 *   much they hold;
 * - captured inline markup is one unit when it shows text and none when it
 *   shows nothing;
 * - bookmark, comment and move boundaries take no units;
 * - hyperlinks, tracked changes, content controls and the transparent
 *   wrappers take none themselves: their content is counted in place.
 *
 * Tracked-deleted text counts like any other: the visible-text view is
 * derived, and positions stay stable when a revision is accepted or rejected
 * elsewhere in the paragraph.
 */

import { panic } from "better-result";

import type {
  Hyperlink,
  InlineSdt,
  Paragraph,
  ParagraphContent,
  Run,
  RunContent,
  TrackedRunContent,
} from "../model/document";

type InlineKind = "run" | "container" | "atom" | "opaque" | "openingMarker" | "closingMarker";

/** How each paragraph child takes part in the offset space. */
const INLINE_KINDS = {
  run: "run",
  hyperlink: "container",
  insertion: "container",
  deletion: "container",
  moveFrom: "container",
  moveTo: "container",
  inlineSdt: "container",
  inlineWrapper: "container",
  simpleField: "atom",
  complexField: "atom",
  mathEquation: "atom",
  commentReference: "atom",
  preservedInline: "opaque",
  bookmarkStart: "openingMarker",
  commentRangeStart: "openingMarker",
  moveFromRangeStart: "openingMarker",
  moveToRangeStart: "openingMarker",
  bookmarkEnd: "closingMarker",
  commentRangeEnd: "closingMarker",
  moveFromRangeEnd: "closingMarker",
  moveToRangeEnd: "closingMarker",
} as const satisfies Record<ParagraphContent["type"], InlineKind>;

type ContainerType = {
  [Type in keyof typeof INLINE_KINDS]: (typeof INLINE_KINDS)[Type] extends "container"
    ? Type
    : never;
}[keyof typeof INLINE_KINDS];

/** A paragraph child whose content is counted in place. */
export type InlineContainer = Extract<ParagraphContent, { type: ContainerType }>;

export const isInlineContainer = (item: ParagraphContent): item is InlineContainer =>
  INLINE_KINDS[item.type] === "container";

/** A zero-width boundary that opens a range: it belongs with what follows it. */
export const isOpeningMarker = (item: ParagraphContent): boolean =>
  INLINE_KINDS[item.type] === "openingMarker";

/** Tracked content that direct text must not join: it is on its way out. */
export const isRemovedRevision = (item: ParagraphContent): boolean =>
  item.type === "deletion" || item.type === "moveFrom";

export const runContentWidth = (content: RunContent): number => {
  switch (content.type) {
    case "text":
      return content.text.length;
    // A cache of where a previous layout broke the page: not content.
    case "renderedPageBreak":
      return 0;
    default:
      return 1;
  }
};

export const runWidth = (run: Run): number => {
  let width = 0;
  for (const content of run.content) {
    width += runContentWidth(content);
  }
  return width;
};

export const childrenOf = (container: InlineContainer): readonly ParagraphContent[] =>
  container.type === "hyperlink" ? container.children : container.content;

export const inlineWidth = (item: ParagraphContent): number => {
  if (item.type === "run") {
    return runWidth(item);
  }
  if (isInlineContainer(item)) {
    return contentWidth(childrenOf(item));
  }
  if (item.type === "preservedInline") {
    return item.text === "" ? 0 : 1;
  }
  return INLINE_KINDS[item.type] === "atom" ? 1 : 0;
};

export const contentWidth = (items: readonly ParagraphContent[]): number => {
  let width = 0;
  for (const item of items) {
    width += inlineWidth(item);
  }
  return width;
};

/** The number of logical units in a paragraph: its largest valid offset. */
export const paragraphLength = (paragraph: Paragraph): number => contentWidth(paragraph.content);

/** The character a unit that is not run text stands for in {@link paragraphLogicalText}. */
export const OBJECT_REPLACEMENT_CHARACTER = "￼";

const appendLogicalText = (items: readonly ParagraphContent[], out: string[]): void => {
  for (const item of items) {
    if (item.type === "run") {
      for (const content of item.content) {
        if (content.type === "text") {
          out.push(content.text);
        } else if (runContentWidth(content) === 1) {
          out.push(OBJECT_REPLACEMENT_CHARACTER);
        }
      }
      continue;
    }
    if (isInlineContainer(item)) {
      appendLogicalText(childrenOf(item), out);
      continue;
    }
    if (inlineWidth(item) === 1) {
      out.push(OBJECT_REPLACEMENT_CHARACTER);
    }
  }
};

/**
 * The paragraph's offset space as a string: run text as it stands and
 * {@link OBJECT_REPLACEMENT_CHARACTER} for every other unit, so index `i` of
 * the string is unit `i` of the paragraph.
 */
export const paragraphLogicalText = (paragraph: Paragraph): string => {
  const out: string[] = [];
  appendLogicalText(paragraph.content, out);
  return out.join("");
};

type HyperlinkChild = Hyperlink["children"][number];
type InlineSdtChild = InlineSdt["content"][number];

const HYPERLINK_CHILD_TYPES = {
  run: true,
  bookmarkStart: true,
  bookmarkEnd: true,
  inlineWrapper: true,
  preservedInline: true,
} as const satisfies Record<HyperlinkChild["type"], true>;

const TRACKED_CHILD_TYPES = {
  run: true,
  hyperlink: true,
  bookmarkStart: true,
  bookmarkEnd: true,
  moveFromRangeStart: true,
  moveFromRangeEnd: true,
  moveToRangeStart: true,
  moveToRangeEnd: true,
  simpleField: true,
  complexField: true,
  inlineSdt: true,
  inlineWrapper: true,
  mathEquation: true,
  preservedInline: true,
  insertion: true,
  deletion: true,
  moveFrom: true,
  moveTo: true,
} as const satisfies Record<TrackedRunContent["type"], true>;

const INLINE_SDT_CHILD_TYPES = {
  run: true,
  hyperlink: true,
  bookmarkStart: true,
  bookmarkEnd: true,
  moveFromRangeStart: true,
  moveFromRangeEnd: true,
  moveToRangeStart: true,
  moveToRangeEnd: true,
  simpleField: true,
  complexField: true,
  inlineSdt: true,
  inlineWrapper: true,
  insertion: true,
  deletion: true,
  moveFrom: true,
  moveTo: true,
  mathEquation: true,
  preservedInline: true,
} as const satisfies Record<InlineSdtChild["type"], true>;

const isHyperlinkChild = (item: ParagraphContent): item is HyperlinkChild =>
  Object.hasOwn(HYPERLINK_CHILD_TYPES, item.type);

const isTrackedChild = (item: ParagraphContent): item is TrackedRunContent =>
  Object.hasOwn(TRACKED_CHILD_TYPES, item.type);

const isInlineSdtChild = (item: ParagraphContent): item is InlineSdtChild =>
  Object.hasOwn(INLINE_SDT_CHILD_TYPES, item.type);

/**
 * Narrow a rebuilt child list back to what its container admits. Every edit
 * keeps the kinds it was given and adds only runs, which each container
 * admits, so a child outside the container's content model is a bug here.
 */
const narrowChildren = <Child extends ParagraphContent>(
  items: readonly ParagraphContent[],
  admits: (item: ParagraphContent) => item is Child,
  container: ContainerType,
): Child[] => {
  const out: Child[] = [];
  for (const item of items) {
    if (!admits(item)) {
      return panic(`A ${item.type} cannot be a child of ${container}.`);
    }
    out.push(item);
  }
  return out;
};

/** The same container holding other children. */
export const withChildren = (
  container: InlineContainer,
  items: readonly ParagraphContent[],
): InlineContainer => {
  switch (container.type) {
    case "hyperlink":
      return { ...container, children: narrowChildren(items, isHyperlinkChild, container.type) };
    case "insertion":
    case "deletion":
    case "moveFrom":
    case "moveTo":
      return { ...container, content: narrowChildren(items, isTrackedChild, container.type) };
    case "inlineSdt":
      return { ...container, content: narrowChildren(items, isInlineSdtChild, container.type) };
    case "inlineWrapper":
      return { ...container, content: [...items] };
    default: {
      const unreachable: never = container;
      return unreachable;
    }
  }
};
