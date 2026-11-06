/**
 * A footnote taller than the room its page leaves continues on the next page.
 *
 * The note's reference line keeps its page and the note fills what is left of
 * it. What remains opens the next page's note area, ahead of that page's own
 * notes, and the body there moves up to make room. A note never breaks through
 * a line: each page shows a window of the note between two line bottoms.
 */

import { measuredLineAdvance } from "./lineFlow";
import type { FootnoteContent, FootnoteWindow, Measure, Page } from "./types";
import {
  FOOTNOTE_ENTRY_MARGIN_BOTTOM,
  FOOTNOTE_SEPARATOR_HEIGHT,
} from "./types";

/** The most of a page a note may claim, keeping body beside its reference. */
const NOTE_SHARE_OF_PAGE = 0.75;
const EPSILON = 0.01;

type PageBox = Pick<Page, "size" | "margins">;

const contentHeightOf = (page: PageBox): number =>
  page.size.h - page.margins.top - page.margins.bottom;

const noteShareOf = (contentHeight: number): number =>
  Math.max(0, contentHeight) * NOTE_SHARE_OF_PAGE;

/** What each note's reference reserves on a page of this content height. */
export const footnoteReservationHeights = (
  heights: ReadonlyMap<number, number>,
  contentHeight: number,
): Map<number, number> => {
  const share = noteShareOf(contentHeight);
  const reserved = new Map<number, number>();
  for (const [id, height] of heights) {
    reserved.set(id, Math.min(height, share));
  }
  return reserved;
};

const blockHeight = (measure: Measure): number => {
  if (measure.kind === "paragraph" || measure.kind === "table") {
    return measure.totalHeight;
  }
  return measure.kind === "image" || measure.kind === "textBox"
    ? measure.height
    : 0;
};

/** Where a block can be cut inside itself: below a line or below a row. */
const inBlockBreakOffsets = (measure: Measure, top: number): number[] => {
  const offsets: number[] = [];
  let bottom = top;
  if (measure.kind === "paragraph") {
    // A note's lines start at its block's top; paragraph spacing trails them.
    for (const line of measure.lines.slice(0, -1)) {
      bottom += measuredLineAdvance(line);
      offsets.push(bottom);
    }
  } else if (measure.kind === "table") {
    for (const row of measure.rows.slice(0, -1)) {
      bottom += row.height;
      offsets.push(bottom);
    }
  }
  return offsets;
};

/** Every height a window of the note may end at, ascending, the note's end last. */
export const footnoteBreakOffsets = (
  content: Pick<FootnoteContent, "measures" | "height">,
): number[] => {
  const offsets: number[] = [];
  let top = 0;
  for (const measure of content.measures) {
    offsets.push(...inBlockBreakOffsets(measure, top));
    top += blockHeight(measure);
    offsets.push(top);
  }
  offsets.push(content.height);
  return [...new Set(offsets)].toSorted((a, b) => a - b);
};

type Entry = { noteId: number; from: number };

type NoteContent = Pick<FootnoteContent, "measures" | "height">;

const bodyBottomOf = (page: Page): number => {
  let bottom = 0;
  for (const fragment of page.fragments) {
    bottom = Math.max(
      bottom,
      fragment.y + ("height" in fragment ? fragment.height : 0),
    );
  }
  return Math.max(0, bottom - page.margins.top);
};

type PlanOptions = {
  pages: readonly Page[];
  contentById: ReadonlyMap<number, NoteContent>;
  /** Pages already added past the last body page for a note's rest. */
  trailingPages?: number;
};

export type FootnoteContinuationPlan = {
  /** A page number's notes and their height, for each page the plan changes. */
  pages: Map<
    number,
    { ids: number[]; windows: FootnoteWindow[]; height: number }
  >;
  /** The height a carried note reserves on a page before the body is placed. */
  reservedHeights: Map<number, number>;
  /** Pages the rest of a note still needs beyond the ones laid out. */
  extraPages: number;
  /** How many pages past the last body page the layout should add. */
  trailingPages: number;
};

type PlaceOptions = {
  entries: readonly Entry[];
  /** What the page leaves for its notes after the body. */
  room: number;
  /** What the same page would leave with no body at all. */
  fullRoom: number;
  contentById: ReadonlyMap<number, NoteContent>;
};

/** What of these notes fits a page's room, and what goes on to the next page. */
const placeEntries = ({
  entries,
  room,
  fullRoom,
  contentById,
}: PlaceOptions) => {
  const ids: number[] = [];
  const windows: FootnoteWindow[] = [];
  const rest: Entry[] = [];
  let left = room;
  let cut = false;
  let height = FOOTNOTE_SEPARATOR_HEIGHT;
  for (const { noteId, from } of entries) {
    const content = contentById.get(noteId);
    if (!content) continue;
    if (cut) {
      rest.push({ noteId, from });
      continue;
    }
    const remaining = content.height - from;
    if (remaining + FOOTNOTE_ENTRY_MARGIN_BOTTOM <= left + EPSILON) {
      ids.push(noteId);
      if (from > 0) windows.push({ noteId, from, to: content.height });
      left -= remaining + FOOTNOTE_ENTRY_MARGIN_BOTTOM;
      height += remaining + FOOTNOTE_ENTRY_MARGIN_BOTTOM;
      continue;
    }
    const start = from + EPSILON;
    const offsets = footnoteBreakOffsets(content).filter(
      (offset) => offset > start,
    );
    let fitting: number | undefined;
    for (const offset of offsets) {
      if (offset - from <= left + EPSILON) fitting = offset;
    }
    // A line taller than any page still has to show somewhere. One that only
    // lacks room here waits for a later page instead of covering the body.
    const next = offsets.at(0);
    const oversized =
      ids.length === 0 &&
      next !== undefined &&
      next - from > fullRoom + EPSILON;
    const to = fitting ?? (oversized ? next : undefined);
    cut = true;
    if (to === undefined) {
      rest.push({ noteId, from });
      continue;
    }
    ids.push(noteId);
    windows.push({ noteId, from, to });
    height += to - from + FOOTNOTE_ENTRY_MARGIN_BOTTOM;
    if (to < content.height - EPSILON) rest.push({ noteId, from: to });
  }
  return { ids, windows, rest, height: ids.length === 0 ? 0 : height };
};

/** Where each note's windows fall, given how the body laid out. */
export const planFootnoteContinuation = ({
  pages,
  contentById,
  trailingPages = 0,
}: PlanOptions): FootnoteContinuationPlan => {
  const plan: FootnoteContinuationPlan = {
    pages: new Map(),
    reservedHeights: new Map(),
    extraPages: 0,
    trailingPages,
  };
  const last = pages.at(-1);
  if (!last) return plan;

  // Every page the rest will reach reserves its share now, so the body does
  // not have to find them one layout at a time.
  const reserveFor = (entries: readonly Entry[], fromPage: number) => {
    let needed = FOOTNOTE_SEPARATOR_HEIGHT;
    for (const { noteId, from } of entries) {
      const height = contentById.get(noteId)?.height ?? from;
      needed += height - from + FOOTNOTE_ENTRY_MARGIN_BOTTOM;
    }
    for (let next = fromPage; needed > EPSILON; next += 1) {
      const share = noteShareOf(contentHeightOf(pages[next - 1] ?? last));
      if (share <= EPSILON) {
        // A page with no content area takes nothing, and none past the end will.
        if (next > pages.length) break;
        continue;
      }
      plan.reservedHeights.set(next, Math.min(needed, share));
      needed -= share;
    }
  };

  let carried: Entry[] = [];
  for (const page of pages) {
    const own = (page.footnoteIds ?? []).map((noteId) => ({ noteId, from: 0 }));
    const fullRoom = contentHeightOf(page) - FOOTNOTE_SEPARATOR_HEIGHT;
    const placed = placeEntries({
      entries: [...carried, ...own],
      room: fullRoom - bodyBottomOf(page),
      fullRoom,
      contentById,
    });
    if (carried.length > 0 || placed.rest.length > 0) {
      plan.pages.set(page.number, {
        ids: placed.ids,
        windows: placed.windows,
        height: placed.height,
      });
    }
    carried = placed.rest;
    if (carried.length > 0) reserveFor(carried, page.number + 1);
  }
  const fullRoom = contentHeightOf(last) - FOOTNOTE_SEPARATOR_HEIGHT;
  while (carried.length > 0) {
    plan.extraPages += 1;
    carried = placeEntries({
      entries: carried,
      room: fullRoom,
      fullRoom,
      contentById,
    }).rest;
  }

  let surplus = 0;
  if (plan.extraPages === 0) {
    const noteOn = (page: Page) =>
      plan.pages.get(page.number)?.ids.length ?? page.footnoteIds?.length ?? 0;
    for (const page of pages.slice(pages.length - trailingPages).toReversed()) {
      if (page.fragments.length > 0 || noteOn(page) > 0) break;
      surplus += 1;
    }
  }
  plan.trailingPages = trailingPages + plan.extraPages - surplus;
  return plan;
};

/** Whether a layout was made with what this plan asks for. */
export const footnoteContinuationSettled = (
  used: {
    reservedHeights: ReadonlyMap<number, number> | undefined;
    trailingPages: number;
  },
  plan: FootnoteContinuationPlan,
): boolean => {
  if (used.trailingPages !== plan.trailingPages) return false;
  if ((used.reservedHeights?.size ?? 0) !== plan.reservedHeights.size) {
    return false;
  }
  for (const [page, height] of plan.reservedHeights) {
    if (Math.abs((used.reservedHeights?.get(page) ?? -1) - height) > EPSILON) {
      return false;
    }
  }
  return true;
};

/** The pages with a plan's note windows. */
export const applyFootnoteContinuation = (
  pages: readonly Page[],
  plan: FootnoteContinuationPlan,
): Page[] =>
  pages.map((page) => {
    const placed = plan.pages.get(page.number);
    if (!placed) return page;
    const next: Page = {
      ...page,
      footnoteIds: placed.ids,
      footnoteReservedHeight: placed.height,
    };
    if (placed.windows.length > 0) next.footnoteWindows = placed.windows;
    return next;
  });
