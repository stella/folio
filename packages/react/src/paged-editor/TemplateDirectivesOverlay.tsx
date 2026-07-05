/**
 * Template Directives Overlay
 *
 * Paints a subtle, translucent highlight over each {{...}} marker's range so a
 * marker reads as a token while staying fully visible and editable: the
 * highlight is faint and `pointer-events: none`, so the real text shows through
 * and the caret lands in it normally. Because it's a tint (not an opaque cover),
 * minor misalignment during reflow is invisible — no flicker. Block directives
 * additionally get a thin left-margin gutter rail spanning opener→closer.
 *
 * Rails are quiet by default and loud on intent: every rail renders faint until
 * the block either contains the caret (innermost wins) or is hovered (via its
 * rail or one of its {{#if}}/{{/if}} chips), at which point that one block's
 * rail and both chips brighten so it can be traced end-to-end. Rails are clipped
 * per page (never across the inter-page gap, header, or footer) and confined to
 * the left margin by a depth budget derived from the margin width.
 *
 * Appearance lives in editor.css (`.folio-template-*`, --doc-* tokens); only
 * positioning is inline.
 */

import type { CSSProperties, PointerEvent as ReactPointerEvent } from "react";
import { useState } from "react";

import type { SelectionRect } from "@stll/folio-core/layout-bridge/engine/selectionRects";
import type {
  DirectiveGutterGeometry,
  PageContentBand,
} from "@stll/folio-core/paged-layout/rangeProjection";
import type {
  DirectiveKind,
  DirectiveRange,
} from "@stll/folio-core/prosemirror/plugins/templateDirectives";

export type DirectiveRectGroup = {
  range: DirectiveRange;
  rects: SelectionRect[];
};

export type TemplateDirectivesOverlayProps = {
  groups: DirectiveRectGroup[];
  /**
   * Left-margin geometry: text-column anchor, usable margin width, and per-page
   * content bands. Null until the pages are measured; rails are suppressed then.
   */
  gutter: DirectiveGutterGeometry | null;
  /**
   * Current caret/selection head PM position, or null. The innermost block
   * containing it is emphasised (the "loud on caret" half of the interaction).
   */
  caretPos: number | null;
  /**
   * Full list of scanned directive ranges in the document. Pairing and nesting
   * depth are derived from this authoritative list (not from `groups`, which is
   * the projected/visible subset), so depth and pairing always reflect what the
   * document actually contains even while pages are off-screen or unprojected.
   */
  ranges: readonly DirectiveRange[];
};

const overlayStyles: CSSProperties = {
  position: "absolute",
  top: 0,
  left: "50%",
  width: "100vw",
  height: "100%",
  transform: "translateX(-50%)",
  pointerEvents: "none",
  zIndex: 10,
};

const BLOCK_OPENERS = new Set<DirectiveKind>(["if", "each"]);
const BLOCK_CLOSERS = new Set<DirectiveKind>(["endif", "endeach"]);

const RAIL_WIDTH = 2.5;
/** Horizontal gap between adjacent nesting depths' rails. */
const RAIL_DEPTH_GAP = 5;
/** Minimum clear space between the innermost (deepest) rail and the text edge. */
const RAIL_TEXT_CLEARANCE = 6;
/** Keep the outermost rail this far off the physical page-left edge. */
const RAIL_EDGE_PAD = 2;
/** Drop per-page segments thinner than this (sub-pixel slivers at a page seam). */
const RAIL_SEGMENT_MIN_HEIGHT = 2;

/** Opener kind of a band ("if" ⇒ condition family, "each" ⇒ loop family). */
type BandKind = "if" | "each";

const bandKindOf = (kind: DirectiveKind): BandKind => (kind === "each" ? "each" : "if");

/**
 * Deepest nesting level whose rail still fits inside the left margin with a
 * `RAIL_TEXT_CLEARANCE` gap to the text and a `RAIL_EDGE_PAD` gap to the page
 * edge. Depths past this reuse the deepest offset, so a 5-level nest can never
 * march a rail into the text column. Pure function of the measured margin width.
 */
export const railBudgetDepth = (marginWidth: number): number => {
  const usable = marginWidth - RAIL_TEXT_CLEARANCE - RAIL_WIDTH - RAIL_EDGE_PAD;
  if (usable <= 0) {
    return 0;
  }
  return Math.floor(usable / RAIL_DEPTH_GAP);
};

/**
 * Left edge (container space) of a depth-`depth` rail. The deepest rail the
 * budget allows sits `RAIL_TEXT_CLEARANCE + RAIL_WIDTH` left of the text edge;
 * shallower (outer) blocks fan further left into the margin — nested blocks read
 * as concentric brackets `[ [ [ text ] ] ]`. Depths beyond the budget clamp to
 * the deepest offset.
 */
export const railXForDepth = (depth: number, contentLeft: number, marginWidth: number): number => {
  const budget = railBudgetDepth(marginWidth);
  const capped = Math.min(Math.max(0, depth), budget);
  const deepestX = contentLeft - RAIL_TEXT_CLEARANCE - RAIL_WIDTH;
  return deepestX - (budget - capped) * RAIL_DEPTH_GAP;
};

export type BandSegment = { top: number; bottom: number };

/**
 * Clip a band's vertical span to each page's body-content band so a rail is
 * drawn as one rounded segment per page and never paints across the inter-page
 * gap, header, or footer. With no page bands (geometry not ready) the whole span
 * is returned as a single segment so rails still show. Pure math.
 */
export const segmentBandByPages = (
  band: BandSegment,
  pageBands: readonly PageContentBand[],
): BandSegment[] => {
  if (pageBands.length === 0) {
    return band.bottom - band.top >= RAIL_SEGMENT_MIN_HEIGHT ? [band] : [];
  }
  const segments: BandSegment[] = [];
  for (const page of pageBands) {
    const top = Math.max(band.top, page.top);
    const bottom = Math.min(band.bottom, page.bottom);
    if (bottom - top >= RAIL_SEGMENT_MIN_HEIGHT) {
      segments.push({ top, bottom });
    }
  }
  return segments;
};

/** A matched block opener→closer pair, derived purely from the scanned ranges. */
export type BlockPairing = {
  /** Stable block id: the opener's `from` PM position. */
  blockId: number;
  kind: BandKind;
  /**
   * 0-based nesting depth of the opener, recorded during the same pairing walk
   * (stack size at push time). Sharing one walk with the pairing guarantees the
   * rail's depth and its opener→closer span can never disagree.
   */
  depth: number;
  openerFrom: number;
  closerFrom: number;
  /** Exclusive PM end of the closer marker (band bottom / caret containment). */
  closerTo: number;
  /** Opener expression, e.g. `hasVerdicts` or `contracts.risks`, for the hint. */
  openerExpr: string;
};

/** One open block awaiting its closer, plus the depth captured when it was pushed. */
type OpenBlock = { range: DirectiveRange; depth: number };

/**
 * Pair block openers with their closers via a kind-aware stack, so both chips of
 * a pair share a `blockId`, the rail knows its opener→closer PM span, and each
 * pairing carries the `depth` recorded in this very walk (never a second,
 * possibly-divergent depth pass). A closer matches the nearest opener of the same
 * family ({{/if}} ⇒ {{#if}}, {{/each}} ⇒ {{#each}}) and drops any still-open
 * openers nested above it; a closer with no matching opener is ignored. This keeps
 * a mid-edit / unbalanced template from pairing a {{/if}} with an {{#each}}.
 * Inline (block:false) markers are excluded. Pure function of the ranges.
 */
export const pairBlockRanges = (ranges: readonly DirectiveRange[]): BlockPairing[] => {
  const blocks = ranges
    .filter((r) => r.block && (BLOCK_OPENERS.has(r.kind) || BLOCK_CLOSERS.has(r.kind)))
    .slice()
    .sort((a, b) => a.from - b.from);

  const pairings: BlockPairing[] = [];
  const stack: OpenBlock[] = [];
  for (const range of blocks) {
    if (BLOCK_OPENERS.has(range.kind)) {
      stack.push({ range, depth: stack.length });
      continue;
    }
    const wantOpener: DirectiveKind = range.kind === "endif" ? "if" : "each";
    let matched: OpenBlock | undefined;
    let matchIdx = -1;
    for (let i = stack.length - 1; i >= 0; i -= 1) {
      const entry = stack[i];
      if (entry?.range.kind === wantOpener) {
        matched = entry;
        matchIdx = i;
        break;
      }
    }
    if (!matched) {
      continue;
    }
    stack.length = matchIdx;
    pairings.push({
      blockId: matched.range.from,
      kind: bandKindOf(matched.range.kind),
      depth: matched.depth,
      openerFrom: matched.range.from,
      closerFrom: range.from,
      closerTo: range.to,
      openerExpr: matched.range.expr,
    });
  }
  return pairings;
};

/** Hover hint for a closer chip: what it closes, e.g. `/if · hasVerdicts`. */
export const closerHintLabel = (kind: BandKind, openerExpr: string): string => {
  const head = kind === "each" ? "/each" : "/if";
  return openerExpr ? `${head} · ${openerExpr}` : head;
};

/**
 * Block id of the innermost pairing whose opener→closer span contains `caretPos`
 * (largest opener position wins), or null. Drives the caret emphasis.
 */
export const innermostBlockAt = (
  pairings: readonly BlockPairing[],
  caretPos: number | null,
): number | null => {
  if (caretPos === null) {
    return null;
  }
  let bestBlockId: number | null = null;
  let bestOpener = Number.NEGATIVE_INFINITY;
  for (const pairing of pairings) {
    const contains = caretPos >= pairing.openerFrom && caretPos <= pairing.closerTo;
    if (contains && pairing.openerFrom > bestOpener) {
      bestOpener = pairing.openerFrom;
      bestBlockId = pairing.blockId;
    }
  }
  return bestBlockId;
};

type RailBand = {
  blockId: number;
  kind: BandKind;
  railX: number;
  segments: BandSegment[];
};

/**
 * Assemble the per-page rail segments for every balanced block: look up
 * opener/closer rects by `from`, place the rail by the pairing's own budgeted
 * depth, and clip its span to each page. Depth comes from the pairing (same walk
 * that matched opener→closer), so a rail's indentation and its span always agree.
 * O(pairings); all geometry comes from the props.
 */
const computeRailBands = (
  pairings: readonly BlockPairing[],
  groups: DirectiveRectGroup[],
  gutter: DirectiveGutterGeometry,
): RailBand[] => {
  const groupByFrom = new Map(groups.map((group) => [group.range.from, group]));

  const bands: RailBand[] = [];
  for (const pairing of pairings) {
    const openerRect = groupByFrom.get(pairing.openerFrom)?.rects[0];
    const closerRect = groupByFrom.get(pairing.closerFrom)?.rects[0];
    if (!openerRect || !closerRect) {
      continue;
    }
    const segments = segmentBandByPages(
      { top: openerRect.y, bottom: closerRect.y + closerRect.height },
      gutter.pageBands,
    );
    if (segments.length === 0) {
      continue;
    }
    bands.push({
      blockId: pairing.blockId,
      kind: pairing.kind,
      railX: railXForDepth(pairing.depth, gutter.contentLeft, gutter.marginWidth),
      segments,
    });
  }
  return bands;
};

const blockIdFromTarget = (target: EventTarget | null): number | null => {
  if (!(target instanceof HTMLElement)) {
    return null;
  }
  const raw = target.dataset["folioBlockId"];
  return raw === undefined ? null : Number(raw);
};

export const TemplateDirectivesOverlay = ({
  groups,
  gutter,
  caretPos,
  ranges,
}: TemplateDirectivesOverlayProps) => {
  const [hoveredBlockId, setHoveredBlockId] = useState<number | null>(null);

  if (groups.length === 0) {
    return null;
  }

  // Pair over the full scanned `ranges` (the authoritative document list), not
  // the projected `groups` subset, so pairing/depth reflect the whole document;
  // `groups` only supplies rect geometry, looked up by `from` in computeRailBands.
  const pairings = pairBlockRanges(ranges);
  const railBands = gutter ? computeRailBands(pairings, groups, gutter) : [];

  // Block id for every opener/closer chip (both share the opener's id) and the
  // closer chips' hover hints, so a chip can pair-highlight and label itself.
  const blockIdByFrom = new Map<number, number>();
  const closerLabelByFrom = new Map<number, string>();
  for (const pairing of pairings) {
    blockIdByFrom.set(pairing.openerFrom, pairing.blockId);
    blockIdByFrom.set(pairing.closerFrom, pairing.blockId);
    closerLabelByFrom.set(pairing.closerFrom, closerHintLabel(pairing.kind, pairing.openerExpr));
  }

  const caretBlockId = innermostBlockAt(pairings, caretPos);
  const isActive = (blockId: number): boolean =>
    blockId === hoveredBlockId || blockId === caretBlockId;

  // Delegated hover: only rails and block chips carry a data-block-id and
  // pointer-events, so one pair of container listeners tracks the hovered block
  // without a closure per element. Staying within the same block keeps it lit.
  const handlePointerOver = (event: ReactPointerEvent) => {
    const blockId = blockIdFromTarget(event.target);
    if (blockId !== null) {
      setHoveredBlockId(blockId);
    }
  };
  const handlePointerOut = (event: ReactPointerEvent) => {
    const from = blockIdFromTarget(event.target);
    if (from === null) {
      return;
    }
    const to = blockIdFromTarget(event.relatedTarget);
    if (to === from) {
      return;
    }
    setHoveredBlockId((current) => (current === from ? to : current));
  };

  return (
    <div
      style={overlayStyles}
      data-folio-template-directives-overlay=""
      onPointerOver={handlePointerOver}
      onPointerOut={handlePointerOut}
    >
      {railBands.flatMap((band) =>
        band.segments.map((segment, segmentIdx) => (
          <div
            key={`rail:${band.blockId}:${segmentIdx}`}
            className={`folio-template-band-rail folio-template-band-rail--${band.kind}${isActive(band.blockId) ? " folio-template-band-rail--active" : ""}`}
            data-folio-block-id={band.blockId}
            style={{
              left: band.railX,
              top: segment.top,
              width: RAIL_WIDTH,
              height: Math.max(0, segment.bottom - segment.top),
            }}
          />
        )),
      )}

      {groups.flatMap(({ range, rects }, groupIdx) => {
        const blockId = range.block ? blockIdByFrom.get(range.from) : undefined;
        const isBlockChip =
          range.block && (BLOCK_OPENERS.has(range.kind) || BLOCK_CLOSERS.has(range.kind));
        const closerLabel = range.block ? closerLabelByFrom.get(range.from) : undefined;
        const active = blockId !== undefined && isActive(blockId);
        return rects.map((rect, idx) => (
          <span
            key={`d:${groupIdx}:${idx}`}
            className={`folio-template-directive folio-template-directive--${range.kind}${isBlockChip ? " folio-template-directive--block" : ""}${active ? " folio-template-directive--active" : ""}`}
            {...(blockId !== undefined ? { "data-folio-block-id": blockId } : {})}
            {...(closerLabel !== undefined ? { title: closerLabel } : {})}
            style={{
              left: rect.x - 1,
              top: rect.y + 1,
              width: rect.width + 2,
              height: Math.max(0, rect.height - 2),
            }}
          />
        ));
      })}
    </div>
  );
};
