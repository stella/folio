/**
 * Building a page's primitives and the regions over them in one walk.
 *
 * A hit region is a box a click can land in, and every one of them is something
 * the producer already knows while it paints: this is a line, these are its
 * runs, that cell is at row 2 column 3. Deriving them afterwards would mean a
 * second walk over the same layout, which is the second implementation this
 * whole design exists to avoid; a region is instead opened around the painting
 * that fills it.
 *
 * A region owns a contiguous slice of the page's primitives, because the
 * producer emits a line's background, its glyphs and its decorations together.
 * That is what lets the tree index into the paint list rather than copy it.
 */

import type {
  DisplayHitRegion,
  DisplayHitRegionKind,
  DisplayHitRegionModel,
  DisplayPrimitive,
  DisplayRect,
} from "../types";

export type RegionDescriptor = {
  readonly kind: DisplayHitRegionKind;
  readonly rect: DisplayRect;
  readonly model?: DisplayHitRegionModel;
};

export type PageComposer = {
  /** Append primitives to the page, in paint order. */
  readonly push: (primitives: readonly DisplayPrimitive[]) => void;
  /**
   * Everything `paint` appends belongs to this region. Regions opened inside
   * `paint` become its children, so the tree is the nesting of the walk.
   */
  readonly region: (descriptor: RegionDescriptor, paint: () => void) => void;
  readonly primitives: () => readonly DisplayPrimitive[];
  readonly regions: () => readonly DisplayHitRegion[];
};

type OpenRegion = {
  readonly descriptor: RegionDescriptor;
  readonly from: number;
  readonly children: DisplayHitRegion[];
};

export const createPageComposer = (): PageComposer => {
  const primitives: DisplayPrimitive[] = [];
  const roots: DisplayHitRegion[] = [];
  const open: OpenRegion[] = [];

  const push = (added: readonly DisplayPrimitive[]): void => {
    primitives.push(...added);
  };

  const region = (descriptor: RegionDescriptor, paint: () => void): void => {
    const entry: OpenRegion = { descriptor, from: primitives.length, children: [] };
    open.push(entry);
    try {
      paint();
    } finally {
      open.pop();
    }
    const closed: DisplayHitRegion = {
      kind: descriptor.kind,
      rect: descriptor.rect,
      from: entry.from,
      to: primitives.length,
      children: entry.children,
      ...(descriptor.model === undefined ? {} : { model: descriptor.model }),
    };
    // A region that painted nothing and contains nothing still bounds a click:
    // an empty line is where a caret goes when a paragraph has no text.
    (open.at(-1)?.children ?? roots).push(closed);
  };

  return {
    push,
    region,
    primitives: () => primitives,
    regions: () => roots,
  };
};
