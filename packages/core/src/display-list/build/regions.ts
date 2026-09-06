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

import type { BlockId } from "../../layout-engine/types";
import type {
  DisplayHitRegion,
  DisplayHitRegionKind,
  DisplayHitRegionModel,
  DisplayPrimitive,
  DisplayRect,
  DisplayStoryRef,
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

/**
 * The region a whole block occupies, wherever it is painted.
 *
 * A body fragment, a picture inside a header, a text box inside a cell: each is
 * one block at one place, and a surface asks the same questions of all of them.
 * Taking the box and the range from the same fragment the painting uses is what
 * keeps the answer the one that was drawn.
 */
type BlockRegionOptions = {
  readonly fragment: {
    readonly blockId: BlockId;
    readonly x: number;
    readonly y: number;
    readonly width: number;
    readonly height: number;
    readonly pmStart?: number;
    readonly pmEnd?: number;
  };
  readonly kind: DisplayHitRegionKind;
  readonly context: { readonly story: DisplayStoryRef };
};

export const blockRegion = ({ fragment, kind, context }: BlockRegionOptions): RegionDescriptor => ({
  kind,
  rect: {
    xPx: fragment.x,
    yPx: fragment.y,
    widthPx: fragment.width,
    heightPx: fragment.height,
  },
  model: {
    blockId: String(fragment.blockId),
    ...(fragment.pmStart === undefined || fragment.pmEnd === undefined
      ? {}
      : { pmRange: { start: fragment.pmStart, end: fragment.pmEnd, story: context.story } }),
  },
});
