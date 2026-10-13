import { OP_STORIES, documentStories, storyBody } from "@stll/docx-core/ops";
import type { OpStory } from "@stll/docx-core/ops";
import type { Node as PMNode } from "prosemirror-model";

import { createTextStreamRange, measurePaintedSpanRange } from "../layout-bridge/dom/textStreamDom";
import { getEditorScrollRoot, scrollEditorElementIntoView } from "../paged-layout/editorScrollRoot";
import { headerFooterToProseDoc, toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { resolveFindMatchRanges } from "../prosemirror/findReplaceSelection";
import type { FindMatchRange } from "../prosemirror/findReplaceSelection";
import type { Document } from "../types/document";
import type { Layout } from "../layout-engine/types";
import { selectSectionHeaderFooterRIds } from "../layout-painter/renderPage";
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
};

/** Resolve each story once; box positions are local to their parent story. */
export const resolveDocumentFindRanges = ({
  document,
  bodyDoc,
  matches,
}: Omit<FindMatchViewOptions, "container" | "commentRoot">): ResolvedMatch[] => {
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
        : headerFooterToProseDoc(storyBody(document, story).content, conversionOptions);
    const groups = new Map<number | undefined, typeof entries>();
    for (const entry of entries) {
      const group = groups.get(entry.match.textBoxIndex);
      if (group) group.push(entry);
      else groups.set(entry.match.textBoxIndex, [entry]);
    }
    const boxes: { doc: PMNode; offset: number }[] = [];
    doc.descendants((node, pos) => {
      if (node.type.name === "textBox") boxes.push({ doc: node, offset: pos + 1 });
      return true;
    });
    for (const [boxIndex, group] of groups) {
      const source = boxIndex === undefined ? { doc, offset: 0 } : boxes.at(boxIndex);
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
  if (story === "header" || story === "footer") return `${story}:${span.dataset["hfRid"]}`;
  if (story === "footnote" || story === "endnote") return `${story}:${span.dataset["noteId"]}`;
  const slot = span.closest<HTMLElement>(
    "[data-hf-slot-kind][data-hf-rid], [data-note-kind][data-note-id]",
  );
  if (slot?.dataset["hfSlotKind"]) return `${slot.dataset["hfSlotKind"]}:${slot.dataset["hfRid"]}`;
  if (slot?.dataset["noteKind"]) return `${slot.dataset["noteKind"]}:${slot.dataset["noteId"]}`;
  return "body";
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
  const commentOverlay = options.matches.some((match) => match.commentId !== undefined)
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
  const targets = new Map<number, HTMLElement>();
  let currentIndex = 0;
  let pendingReveal: number | null = null;
  let frame: number | null = null;
  const window = container.ownerDocument.defaultView;

  const emphasize = () => {
    for (const [index, elements] of marks) {
      for (const element of elements) {
        const active = index === currentIndex;
        element.dataset["current"] = String(active);
        element.style.background = active ? "var(--doc-find-current)" : "var(--doc-find-highlight)";
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
    for (const group of spans.values()) group.sort((a, b) => spanEnd(a) - spanEnd(b));
    const suffixStarts = new Map<string, number[]>();
    for (const [key, group] of spans) {
      const starts = Array.from({ length: group.length }, () => Number.POSITIVE_INFINITY);
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
    const appendRects = ({ index, rects, source, surface }: AppendRectsOptions) => {
      const origin = surface.getBoundingClientRect();
      const scaleX = surface.offsetWidth > 0 ? origin.width / surface.offsetWidth : 1;
      const scaleY = surface.offsetHeight > 0 ? origin.height / surface.offsetHeight : 1;
      const elements = marks.get(index) ?? [];
      const viewport = getEditorScrollRoot(source)?.getBoundingClientRect();
      for (const rect of rects) {
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
          (suffixStarts.get(storyKey(range.story))?.at(index) ?? Number.POSITIVE_INFINITY) >=
          range.to
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
              Math.min(span.textContent?.length ?? 0, atom.endOffset - textOffset),
            )
          : null;
        const measured = atomRange
          ? ({ type: "text", rects: [...atomRange.getClientRects()] } as const)
          : measurePaintedSpanRange(span, range.from, range.to);
        if (measured.type === "outside") continue;
        const previousTarget = targets.get(range.index);
        if (!previousTarget || (previousTarget.compareDocumentPosition(span) & 2) !== 0)
          targets.set(range.index, span);
        const rects = measured.type === "element" ? [measured.rect] : measured.rects;
        appendRects({ index: range.index, rects, source: span, surface: overlay });
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
          Number(element.dataset["findCommentParagraph"]) === match.paragraphIndex,
      );
      if (!paragraph) continue;
      const range = createTextStreamRange(paragraph, match.startOffset, match.endOffset);
      if (!range) continue;
      targets.set(index, paragraph);
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
        scrollEditorElementIntoView(target);
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
  observer.observe(container, { childList: true, subtree: true, attributes: true });
  if (options.commentRoot && options.commentRoot !== container)
    observer.observe(options.commentRoot, { childList: true, subtree: true, attributes: true });
  const resize = new ResizeObserver(schedule);
  resize.observe(container);
  container.addEventListener("scroll", schedule, true);
  options.commentRoot?.addEventListener("scroll", schedule, true);
  repaint();
  return {
    paint: (index: number) => {
      currentIndex = index;
      emphasize();
    },
    reveal: (index: number): boolean => {
      if (!targets.has(index)) repaint();
      const target = targets.get(index);
      pendingReveal = target ? null : index;
      if (!target) return false;
      scrollEditorElementIntoView(target);
      return true;
    },
    destroy: () => {
      observer.disconnect();
      resize.disconnect();
      container.removeEventListener("scroll", schedule, true);
      options.commentRoot?.removeEventListener("scroll", schedule, true);
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
