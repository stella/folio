/**
 * A footnote taller than the room its page leaves continues on the next page.
 *
 * The note's reference line keeps its page and the note fills what is left of
 * it. What remains opens the next page's note area, ahead of that page's own
 * notes, and the body there moves up to make room. A note never breaks through
 * a line: each page shows a window of the note between two line bottoms.
 */

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

/** The height a note's reference reserves on its page while the body is placed. */
export const footnoteReservationHeight = (
  height: number,
  page: PageBox,
): number => Math.min(height, contentHeightOf(page) * NOTE_SHARE_OF_PAGE);

const blockHeight = (measure: Measure): number =>
  measure.kind === "paragraph" || measure.kind === "table"
    ? measure.totalHeight
    : 0;

/** Every height a window of the note may end at, ascending, the note's end last. */
export const footnoteBreakOffsets = (
  content: Pick<FootnoteContent, "measures" | "height">,
): number[] => {
  const offsets: number[] = [];
  let top = 0;
  for (const measure of content.measures) {
    const height = blockHeight(measure);
    if (measure.kind === "paragraph" && measure.lines.length > 0) {
      let lines = 0;
      for (const { lineHeight } of measure.lines) lines += lineHeight;
      let bottom = top + height - lines;
      for (const { lineHeight } of measure.lines.slice(0, -1)) {
        bottom += lineHeight;
        offsets.push(bottom);
      }
    }
    top += height;
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
};

export type FootnoteContinuationPlan = {
  /** A page number's notes and their height, for each page the plan changes. */
  pages: Map<
    number,
    { ids: number[]; windows: FootnoteWindow[]; height: number }
  >;
  /** The height a carried note reserves on a page before the body is placed. */
  reservedHeights: Map<number, number>;
  /** Pages past the last body page that only hold a note's rest. */
  extraPages: number;
};

type PlaceOptions = {
  entries: readonly Entry[];
  room: number;
  contentById: ReadonlyMap<number, NoteContent>;
};

/** What of these notes fits a page's room, and what goes on to the next page. */
const placeEntries = ({ entries, room, contentById }: PlaceOptions) => {
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
    // A line taller than any room still has to show somewhere.
    const to = fitting ?? (ids.length === 0 ? offsets.at(0) : undefined);
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
  return { ids, windows, rest, height };
};

/** Where each note's windows fall, given how the body laid out. */
export const planFootnoteContinuation = ({
  pages,
  contentById,
}: PlanOptions): FootnoteContinuationPlan => {
  const plan: FootnoteContinuationPlan = {
    pages: new Map(),
    reservedHeights: new Map(),
    extraPages: 0,
  };
  const last = pages.at(-1);
  if (!last) return plan;

  let carried: Entry[] = [];
  for (const page of pages) {
    const own = (page.footnoteIds ?? []).map((noteId) => ({ noteId, from: 0 }));
    const placed = placeEntries({
      entries: [...carried, ...own],
      room:
        contentHeightOf(page) - bodyBottomOf(page) - FOOTNOTE_SEPARATOR_HEIGHT,
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
    if (carried.length === 0) continue;
    // Every page the rest will reach reserves its share now, so the body does
    // not have to find them one layout at a time.
    let needed = FOOTNOTE_SEPARATOR_HEIGHT;
    for (const { noteId, from } of carried) {
      const height = contentById.get(noteId)?.height ?? from;
      needed += height - from + FOOTNOTE_ENTRY_MARGIN_BOTTOM;
    }
    const share = contentHeightOf(page) * NOTE_SHARE_OF_PAGE;
    for (let next = page.number + 1; needed > EPSILON; next += 1) {
      plan.reservedHeights.set(next, Math.min(needed, share));
      needed -= share;
    }
  }
  while (carried.length > 0) {
    plan.extraPages += 1;
    const placed = placeEntries({
      entries: carried,
      room: contentHeightOf(last) - FOOTNOTE_SEPARATOR_HEIGHT,
      contentById,
    });
    plan.pages.set(last.number + plan.extraPages, {
      ids: placed.ids,
      windows: placed.windows,
      height: placed.height,
    });
    carried = placed.rest;
  }
  return plan;
};

/** Whether a plan's reserved heights are the ones the layout was made with. */
export const footnoteReservationsSettled = (
  used: ReadonlyMap<number, number> | undefined,
  plan: FootnoteContinuationPlan,
): boolean => {
  if ((used?.size ?? 0) !== plan.reservedHeights.size) return false;
  for (const [page, height] of plan.reservedHeights) {
    if (Math.abs((used?.get(page) ?? -1) - height) > EPSILON) return false;
  }
  return true;
};

/** The pages with a plan's note windows, and any pages only a note's rest needs. */
export const applyFootnoteContinuation = (
  pages: readonly Page[],
  plan: FootnoteContinuationPlan,
): Page[] => {
  const out = pages.map((page) => {
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
  const last = pages.at(-1);
  if (!last) return out;
  for (let extra = 1; extra <= plan.extraPages; extra += 1) {
    const number = last.number + extra;
    const placed = plan.pages.get(number);
    const { endnoteIds: _endnotes, footnoteWindows: _windows, ...shell } = last;
    out.push({
      ...shell,
      number,
      fragments: [],
      footnoteIds: placed?.ids ?? [],
      footnoteReservedHeight: placed?.height ?? 0,
      ...(placed && placed.windows.length > 0
        ? { footnoteWindows: placed.windows }
        : {}),
    });
  }
  return out;
};
