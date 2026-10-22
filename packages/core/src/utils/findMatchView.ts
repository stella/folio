import type { Node as PMNode } from "prosemirror-model";

import { documentStories, OP_STORIES, storyBody } from "@stll/docx-core/ops";
import type { OpStory } from "@stll/docx-core/ops";

import {
  createTextStreamRange,
  measurePaintedSpanRange,
} from "../layout-bridge/dom/textStreamDom";
import type { Layout } from "../layout-engine/types";
import { selectSectionHeaderFooterRIds } from "../layout-painter/renderPage";
import {
  getEditorScrollRoot,
  scrollEditorElementIntoView,
} from "../paged-layout/editorScrollRoot";
import {
  headerFooterToProseDoc,
  toProseDoc,
} from "../prosemirror/conversion/toProseDoc";
import { resolveFindMatchRanges } from "../prosemirror/findReplaceSelection";
import type { FindMatchRange } from "../prosemirror/findReplaceSelection";
import type { Document } from "../types/document";
import type { FindMatch } from "./findReplace";

const storyKey = (story: OpStory): string => {
  if (story === OP_STORIES.MAIN) return "body";
  switch (story.kind) {
    case "header":
    case "footer":
      return `${story.kind}:${story.rId}`;
    case "footnote":
    case "endnote":
      return `${story.kind}:${story.id}`;
    default: {
      const exhaustive: never = story;
      return exhaustive;
    }
  }
};

/** The same page variant selection used by the painter, including unmounted pages. */
export const getRenderedSearchStories = (layout: Layout | null): OpStory[] => {
  const stories = new Map<string, OpStory>();
  for (const page of layout?.pages ?? []) {
    const { headerRId, footerRId } = selectSectionHeaderFooterRIds(page);
    if (headerRId !== undefined)
      stories.set(`header:${headerRId}`, { kind: "header", rId: headerRId });
    if (footerRId !== undefined)
      stories.set(`footer:${footerRId}`, { kind: "footer", rId: footerRId });
  }
  return [...stories.values()];
};

type ResolvedMatch = FindMatchRange & { index: number; story: OpStory };
type FindMatchViewOptions = {
  container: HTMLElement;
  document: Document;
  bodyDoc?: PMNode | undefined;
  matches: readonly FindMatch[];
  commentRoot?: HTMLElement | undefined;
  layout?: Layout | null | undefined;
};

/** Resolve each story once; box positions are local to their parent story. */
export const resolveDocumentFindRanges = ({
  document,
  bodyDoc,
  matches,
}: Omit<
  FindMatchViewOptions,
  "container" | "commentRoot" | "layout"
>): ResolvedMatch[] => {
  const resolved: ResolvedMatch[] = [];
  const { styles, theme, numbering } = document.package;
  const conversionOptions = {
    ...(styles !== undefined && { styles }),
    ...(theme !== undefined && { theme }),
    ...(numbering !== undefined && { numbering }),
  };
  const byStory = new Map<string, { match: FindMatch; index: number }[]>();
  for (const [index, match] of matches.entries()) {
    if (match.commentId !== undefined) continue;
    // A review view's body match already addresses the authored editor document.
    if (match.authoredRange && match.textBoxIndex === undefined) {
      resolved.push({ ...match.authoredRange, index, story: OP_STORIES.MAIN });
      continue;
    }
    const key = storyKey(match.story ?? OP_STORIES.MAIN);
    const group = byStory.get(key);
    if (group) group.push({ match, index });
    else byStory.set(key, [{ match, index }]);
  }
  for (const story of documentStories(document)) {
    const entries = byStory.get(storyKey(story)) ?? [];
    if (entries.length === 0) continue;
    const doc =
      story === OP_STORIES.MAIN
        ? (bodyDoc ?? toProseDoc(document))
        : headerFooterToProseDoc(
            storyBody(document, story).content,
            conversionOptions,
          );
    const groups = new Map<number | undefined, typeof entries>();
    for (const entry of entries) {
      const group = groups.get(entry.match.textBoxIndex);
      if (group) group.push(entry);
      else groups.set(entry.match.textBoxIndex, [entry]);
    }
    const boxes: { doc: PMNode; offset: number }[] = [];
    doc.descendants((node, pos) => {
      if (node.type.name === "textBox")
        boxes.push({ doc: node, offset: pos + 1 });
      return true;
    });
    for (const [boxIndex, group] of groups) {
      const source =
        boxIndex === undefined ? { doc, offset: 0 } : boxes.at(boxIndex);
      if (!source) continue;
      const ranges = resolveFindMatchRanges(
        source.doc,
        group.map(({ match }) => match),
      );
      for (const [index, range] of ranges.entries()) {
        const entry = group.at(index);
        if (range && entry)
          resolved.push({
            from: range.from + source.offset,
            to: range.to + source.offset,
            ...(range.excludedRanges && {
              excludedRanges: range.excludedRanges.map((anchor) => ({
                from: anchor.from + source.offset,
                to: anchor.to + source.offset,
              })),
            }),
            ...(range.atomicTextRanges && {
              atomicTextRanges: range.atomicTextRanges.map((atom) => ({
                from: atom.from + source.offset,
                to: atom.to + source.offset,
                startOffset: atom.startOffset,
                endOffset: atom.endOffset,
              })),
            }),
            index: entry.index,
            story,
          });
      }
    }
  }
  return resolved;
};

const spanStoryKey = (span: HTMLElement): string => {
  const story = span.dataset["story"];
  if (story === "header" || story === "footer")
    return `${story}:${span.dataset["hfRid"]}`;
  if (story === "footnote" || story === "endnote")
    return `${story}:${span.dataset["noteId"]}`;
  const slot = span.closest<HTMLElement>(
    "[data-hf-slot-kind][data-hf-rid], [data-note-kind][data-note-id]",
  );
  if (slot?.dataset["hfSlotKind"])
    return `${slot.dataset["hfSlotKind"]}:${slot.dataset["hfRid"]}`;
  if (slot?.dataset["noteKind"])
    return `${slot.dataset["noteKind"]}:${slot.dataset["noteId"]}`;
  return "body";
};

/** CSS clipping can hide a nonzero range; scroll roots stay revealable off-screen. */
const clipPaintedRects = (source: HTMLElement, rects: readonly DOMRect[]) => {
  const root = getEditorScrollRoot(source);
  const clips: { box: DOMRect; x: boolean; y: boolean }[] = [];
  for (
    let parent = source.parentElement;
    parent && parent !== root;
    parent = parent.parentElement
  ) {
    const style = getComputedStyle(parent);
    const x = style.overflowX === "hidden" || style.overflowX === "clip";
    const y = style.overflowY === "hidden" || style.overflowY === "clip";
    if (x || y) clips.push({ box: parent.getBoundingClientRect(), x, y });
  }
  return rects.flatMap((rect) => {
    let { left, right, top, bottom } = rect;
    for (const { box, x, y } of clips) {
      if (x) {
        left = Math.max(left, box.left);
        right = Math.min(right, box.right);
      }
      if (y) {
        top = Math.max(top, box.top);
        bottom = Math.min(bottom, box.bottom);
      }
    }
    return right > left && bottom > top
      ? [new DOMRect(left, top, right - left, bottom - top)]
      : [];
  });
};

/** Paint search geometry over the existing page DOM, without changing editor state. */
export const createFindMatchView = (options: FindMatchViewOptions) => {
  const { container } = options;
  const ranges = resolveDocumentFindRanges(options);
  const overlay = container.ownerDocument.createElement("div");
  overlay.dataset["folioFindOverlay"] = "";
  overlay.setAttribute("aria-hidden", "true");
  Object.assign(overlay.style, {
    position: "absolute",
    inset: "0",
    pointerEvents: "none",
    zIndex: "2",
  });
  container.append(overlay);
  const commentRoot = options.commentRoot ?? container;
  const commentOverlay = options.matches.some(
    (match) => match.commentId !== undefined,
  )
    ? container.ownerDocument.createElement("div")
    : null;
  if (commentOverlay) {
    commentOverlay.dataset["folioFindCommentOverlay"] = "";
    commentOverlay.setAttribute("aria-hidden", "true");
    Object.assign(commentOverlay.style, {
      position: "absolute",
      inset: "0",
      pointerEvents: "none",
      zIndex: "110",
    });
    commentRoot.append(commentOverlay);
  }
  const marks = new Map<number, HTMLElement[]>();
  const targets = new Map<number, { source: HTMLElement; rect: DOMRect }>();
  let currentIndex = 0;
  let pendingReveal: number | null = null;
  let frame: number | null = null;
  const window = container.ownerDocument.defaultView;

  const emphasize = () => {
    for (const [index, elements] of marks) {
      for (const element of elements) {
        const active = index === currentIndex;
        element.dataset["current"] = String(active);
        element.style.background = active
          ? "var(--doc-find-current)"
          : "var(--doc-find-highlight)";
        element.style.outlineWidth = active ? "2px" : "0px";
        element.style.outlineStyle = "solid";
        element.style.outlineColor = "var(--doc-find-current)";
      }
    }
  };
  const repaint = () => {
    marks.clear();
    targets.clear();
    const spans = new Map<string, HTMLElement[]>();
    for (const span of container.querySelectorAll<HTMLElement>(
      ".layout-page span[data-pm-start][data-pm-end], .layout-page span[data-decoration-start][data-decoration-end]",
    )) {
      const key = spanStoryKey(span);
      const group = spans.get(key);
      if (group) group.push(span);
      else spans.set(key, [span]);
    }
    for (const group of spans.values())
      group.sort((a, b) => spanEnd(a) - spanEnd(b));
    const suffixStarts = new Map<string, number[]>();
    for (const [key, group] of spans) {
      const starts = Array.from(
        { length: group.length },
        () => Number.POSITIVE_INFINITY,
      );
      let minimum = Number.POSITIVE_INFINITY;
      for (let index = group.length - 1; index >= 0; index--) {
        const span = group.at(index);
        if (span) minimum = Math.min(minimum, spanStart(span));
        starts[index] = minimum;
      }
      suffixStarts.set(key, starts);
    }
    const fragment = container.ownerDocument.createDocumentFragment();
    const commentFragment = container.ownerDocument.createDocumentFragment();
    type AppendRectsOptions = {
      index: number;
      rects: readonly DOMRect[];
      source: HTMLElement;
      surface: HTMLElement;
    };
    const appendRects = ({
      index,
      rects,
      source,
      surface,
    }: AppendRectsOptions) => {
      const visibleRects = clipPaintedRects(source, rects);
      if (visibleRects.length === 0) return;
      const previousTarget = targets.get(index);
      const firstRect = visibleRects.at(0);
      if (
        firstRect &&
        (!previousTarget ||
          (previousTarget.source.compareDocumentPosition(source) & 2) !== 0)
      )
        targets.set(index, { source, rect: firstRect });
      const origin = surface.getBoundingClientRect();
      const scaleX =
        surface.offsetWidth > 0 ? origin.width / surface.offsetWidth : 1;
      const scaleY =
        surface.offsetHeight > 0 ? origin.height / surface.offsetHeight : 1;
      const elements = marks.get(index) ?? [];
      const viewport = getEditorScrollRoot(source)?.getBoundingClientRect();
      for (const rect of visibleRects) {
        const mark = container.ownerDocument.createElement("span");
        mark.dataset["findMatchIndex"] = String(index);
        Object.assign(mark.style, {
          position: "absolute",
          left: `${(rect.left - origin.left) / scaleX}px`,
          top: `${(rect.top - origin.top) / scaleY}px`,
          width: `${rect.width / scaleX}px`,
          height: `${rect.height / scaleY}px`,
          borderRadius: "2px",
          ...(viewport && {
            clipPath: `inset(${Math.max(0, viewport.top - rect.top) / scaleY}px ${Math.max(0, rect.right - viewport.right) / scaleX}px ${Math.max(0, rect.bottom - viewport.bottom) / scaleY}px ${Math.max(0, viewport.left - rect.left) / scaleX}px)`,
          }),
        });
        elements.push(mark);
        (surface === overlay ? fragment : commentFragment).append(mark);
      }
      marks.set(index, elements);
    };
    for (const range of ranges) {
      const group = spans.get(storyKey(range.story)) ?? [];
      // Binary search skips runs before the match, including repeated part occurrences.
      let low = 0;
      let high = group.length;
      while (low < high) {
        const middle = (low + high) >>> 1;
        const span = group.at(middle);
        if (span && spanEnd(span) <= range.from) low = middle + 1;
        else high = middle;
      }
      for (let index = low; index < group.length; index++) {
        if (
          (suffixStarts.get(storyKey(range.story))?.at(index) ??
            Number.POSITIVE_INFINITY) >= range.to
        )
          break;
        const span = group.at(index);
        if (!span) continue;
        const atom = range.atomicTextRanges?.find(
          ({ from, to }) => from === spanStart(span) && to === spanEnd(span),
        );
        const textOffset = Number(span.dataset["decorationTextOffset"] ?? 0);
        if (
          atom &&
          (atom.endOffset <= textOffset ||
            atom.startOffset >= textOffset + (span.textContent?.length ?? 0))
        )
          continue;
        const atomRange = atom
          ? createTextStreamRange(
              span,
              Math.max(0, atom.startOffset - textOffset),
              Math.min(
                span.textContent?.length ?? 0,
                atom.endOffset - textOffset,
              ),
            )
          : null;
        let intervals = [{ from: range.from, to: range.to }];
        for (const anchor of range.excludedRanges ?? []) {
          intervals = intervals.flatMap((interval) => {
            if (anchor.to <= interval.from || anchor.from >= interval.to)
              return [interval];
            const retained: { from: number; to: number }[] = [];
            if (interval.from < anchor.from)
              retained.push({ from: interval.from, to: anchor.from });
            if (anchor.to < interval.to)
              retained.push({ from: anchor.to, to: interval.to });
            return retained;
          });
        }
        const measurements = atomRange
          ? [{ type: "text", rects: [...atomRange.getClientRects()] } as const]
          : intervals.map(({ from, to }) =>
              measurePaintedSpanRange(span, from, to),
            );
        for (const measured of measurements) {
          if (measured.type === "outside") continue;
          const rects =
            measured.type === "element" ? [measured.rect] : measured.rects;
          appendRects({
            index: range.index,
            rects,
            source: span,
            surface: overlay,
          });
        }
      }
    }
    const commentParagraphs = [
      ...commentRoot.querySelectorAll<HTMLElement>(
        "[data-find-comment-id][data-find-comment-paragraph]",
      ),
    ];
    for (const [index, match] of options.matches.entries()) {
      if (match.commentId === undefined) continue;
      const paragraph = commentParagraphs.find(
        (element) =>
          Number(element.dataset["findCommentId"]) === match.commentId &&
          Number(element.dataset["findCommentParagraph"]) ===
            match.paragraphIndex,
      );
      if (!paragraph) continue;
      const range = createTextStreamRange(
        paragraph,
        match.startOffset,
        match.endOffset,
      );
      if (!range) continue;
      if (commentOverlay)
        appendRects({
          index,
          rects: [...range.getClientRects()],
          source: paragraph,
          surface: commentOverlay,
        });
    }
    overlay.replaceChildren(fragment);
    commentOverlay?.replaceChildren(commentFragment);
    emphasize();
    if (pendingReveal !== null) {
      const target = targets.get(pendingReveal);
      if (target) {
        pendingReveal = null;
        scrollEditorElementIntoView(target.source, { rect: target.rect });
      }
    }
  };
  const schedule = () => {
    if (frame !== null || !window) return;
    frame = window.requestAnimationFrame(() => {
      frame = null;
      repaint();
    });
  };
  const observer = new MutationObserver((mutations) => {
    if (
      mutations.some(
        ({ target }) =>
          target !== overlay &&
          !overlay.contains(target) &&
          target !== commentOverlay &&
          !commentOverlay?.contains(target),
      )
    )
      schedule();
  });
  observer.observe(container, {
    childList: true,
    subtree: true,
    attributes: true,
  });
  if (options.commentRoot && options.commentRoot !== container)
    observer.observe(options.commentRoot, {
      childList: true,
      subtree: true,
      attributes: true,
    });
  const resize = new ResizeObserver(schedule);
  resize.observe(container);
  // Clip paths follow each source's editor scroll root, which may enclose a
  // surface; a capture listener on the surface alone misses that root's scroll.
  const scrollTargets = new Set<HTMLElement>(
    [container, options.commentRoot].flatMap((surface) =>
      surface ? [surface, getEditorScrollRoot(surface) ?? surface] : [],
    ),
  );
  for (const target of scrollTargets)
    target.addEventListener("scroll", schedule, true);
  repaint();
  const paint = (index: number) => {
    currentIndex = index;
    if (pendingReveal !== index) pendingReveal = null;
    emphasize();
  };
  const reveal = (index: number): boolean => {
    // Thread activation and virtualization may have changed visibility since paint.
    pendingReveal = null;
    repaint();
    const target = targets.get(index);
    pendingReveal = target ? null : index;
    if (!target) return false;
    scrollEditorElementIntoView(target.source, { rect: target.rect });
    return true;
  };
  return {
    paint,
    reveal,
    navigate: (
      index: number,
      activateComment?: (commentId: number) => void,
    ): void => {
      if (!Number.isInteger(index) || index < 0) return;
      const match = options.matches.at(index);
      if (!match) return;
      paint(index);
      if (match.commentId !== undefined) activateComment?.(match.commentId);
      if (reveal(index) || match.commentId !== undefined) return;
      const range = ranges.find((candidate) => candidate.index === index);
      if (!range) return;
      const page = options.layout?.pages.find((candidate) => {
        const story = range.story;
        if (story === OP_STORIES.MAIN)
          return candidate.fragments.some(
            ({ pmStart, pmEnd }) =>
              pmStart !== undefined &&
              pmEnd !== undefined &&
              pmStart < range.to &&
              pmEnd > range.from,
          );
        switch (story.kind) {
          case "header":
          case "footer": {
            const refs = selectSectionHeaderFooterRIds(candidate);
            return (
              (story.kind === "header" ? refs.headerRId : refs.footerRId) ===
              story.rId
            );
          }
          case "footnote":
            return candidate.footnoteIds?.includes(story.id) ?? false;
          case "endnote":
            return candidate.endnoteIds?.includes(story.id) ?? false;
          default: {
            const exhaustive: never = story;
            return exhaustive;
          }
        }
      });
      if (!page) return;
      // Shells retain their dimensions while page content is unmounted. Scrolling
      // mounts the page; the geometry observer finishes the pending reveal.
      scrollEditorElementIntoView(
        container.querySelector<HTMLElement>(
          `[data-page-number="${page.number}"]`,
        ),
      );
    },
    destroy: () => {
      observer.disconnect();
      resize.disconnect();
      for (const target of scrollTargets)
        target.removeEventListener("scroll", schedule, true);
      if (frame !== null) window?.cancelAnimationFrame(frame);
      overlay.remove();
      commentOverlay?.remove();
    },
  };
};

const spanStart = (span: HTMLElement) =>
  Number(span.dataset["pmStart"] ?? span.dataset["decorationStart"]);
const spanEnd = (span: HTMLElement) =>
  Number(span.dataset["pmEnd"] ?? span.dataset["decorationEnd"]);
