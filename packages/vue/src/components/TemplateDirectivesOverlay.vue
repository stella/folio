<!--
  Template Directives Overlay (Vue)

  Paints a subtle translucent highlight over each {{...}} marker's range so a
  marker reads as a token while staying visible and editable. Block directives
  additionally get a thin left-margin gutter rail spanning opener→closer, quiet
  by default and loud on caret/hover intent.

  Port of the React `TemplateDirectivesOverlay`. Directive ranges + caret come
  from the live editor state (via the `templateDirectives` plugin); rects are
  projected with core's `projectRangesToRects` and the gutter geometry with
  `measureDirectiveGutter`. Both return unscaled (page-space) coordinates, so
  every value is multiplied back by `zoom` at render time (the Vue pages carry
  `transform: scale(zoom)`; this overlay does not).
-->
<template>
  <div
    class="folio-template-directives-overlay"
    data-folio-template-directives-overlay=""
    @pointerover="handlePointerOver"
    @pointerout="handlePointerOut"
  >
    <template v-for="band in railBands" :key="`rail:${band.blockId}`">
      <div
        v-for="(segment, segmentIdx) in band.segments"
        :key="`rail:${band.blockId}:${segmentIdx}`"
        :class="[
          'folio-template-band-rail',
          `folio-template-band-rail--${band.kind}`,
          { 'folio-template-band-rail--active': isActive(band.blockId) },
        ]"
        :data-folio-block-id="band.blockId"
        :style="{
          position: 'absolute',
          left: `${band.railX * zoom}px`,
          top: `${segment.top * zoom}px`,
          width: `${RAIL_WIDTH * zoom}px`,
          height: `${Math.max(0, segment.bottom - segment.top) * zoom}px`,
        }"
      />
    </template>

    <template v-for="(group, groupIdx) in projectedGroups" :key="`d:${groupIdx}`">
      <span
        v-for="(rect, idx) in group.rects"
        :key="`d:${groupIdx}:${idx}`"
        :class="chipClass(group.range)"
        :data-folio-block-id="blockIdFor(group.range)"
        :title="closerTitleFor(group.range)"
        :style="{
          position: 'absolute',
          left: `${(rect.x - 1) * zoom}px`,
          top: `${(rect.y + 1) * zoom}px`,
          width: `${(rect.width + 2) * zoom}px`,
          height: `${Math.max(0, rect.height - 2) * zoom}px`,
        }"
      />
    </template>
  </div>
</template>

<script setup lang="ts">
import { computed, onBeforeUnmount, onMounted, ref, watch } from "vue";
import type { EditorState } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";

import type { SelectionRect } from "@stll/folio-core/layout-bridge/engine/selectionRects";
import type { FlowBlock, Layout, Measure } from "@stll/folio-core/layout-engine/types";
import type {
  DirectiveGutterGeometry,
  PageContentBand,
} from "@stll/folio-core/paged-layout/rangeProjection";
import {
  measureDirectiveGutter,
  projectRangesToRects,
} from "@stll/folio-core/paged-layout/rangeProjection";
import type {
  DirectiveKind,
  DirectiveRange,
} from "@stll/folio-core/prosemirror/plugins/templateDirectives";
import { getTemplateDirectives } from "@stll/folio-core/prosemirror/plugins/templateDirectives";

const props = defineProps<{
  getView: () => EditorView | null;
  getPagesContainer: () => HTMLElement | null;
  editorState: EditorState | null;
  zoom: number;
  layout: Layout | null;
  blocks: FlowBlock[];
  measures: Measure[];
}>();

// ---- Rail geometry (ported verbatim from the React overlay) --------------
const RAIL_WIDTH = 2.5;
/** Horizontal gap between adjacent nesting depths' rails. */
const RAIL_DEPTH_GAP = 5;
/** Minimum clear space between the innermost (deepest) rail and the text edge. */
const RAIL_TEXT_CLEARANCE = 6;
/** Keep the outermost rail this far off the physical page-left edge. */
const RAIL_EDGE_PAD = 2;
/** Drop per-page segments thinner than this (sub-pixel slivers at a page seam). */
const RAIL_SEGMENT_MIN_HEIGHT = 2;

const BLOCK_OPENERS = new Set<DirectiveKind>(["if", "each"]);
const BLOCK_CLOSERS = new Set<DirectiveKind>(["endif", "endeach"]);

type BandKind = "if" | "each";
type BandSegment = { top: number; bottom: number };
type BlockPairing = {
  blockId: number;
  kind: BandKind;
  depth: number;
  openerFrom: number;
  closerFrom: number;
  closerTo: number;
  openerExpr: string;
};
type OpenBlock = { range: DirectiveRange; depth: number };
type RailBand = {
  blockId: number;
  kind: BandKind;
  railX: number;
  segments: BandSegment[];
};
type ProjectedDirectiveGroup = { range: DirectiveRange; rects: SelectionRect[] };

const bandKindOf = (kind: DirectiveKind): BandKind => (kind === "each" ? "each" : "if");

function railBudgetDepth(marginWidth: number): number {
  const usable = marginWidth - RAIL_TEXT_CLEARANCE - RAIL_WIDTH - RAIL_EDGE_PAD;
  if (usable <= 0) {
    return 0;
  }
  return Math.floor(usable / RAIL_DEPTH_GAP);
}

function railXForDepth(depth: number, contentLeft: number, marginWidth: number): number {
  const budget = railBudgetDepth(marginWidth);
  const capped = Math.min(Math.max(0, depth), budget);
  const deepestX = contentLeft - RAIL_TEXT_CLEARANCE - RAIL_WIDTH;
  return deepestX - (budget - capped) * RAIL_DEPTH_GAP;
}

function segmentBandByPages(
  band: BandSegment,
  pageBands: readonly PageContentBand[],
): BandSegment[] {
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
}

function pairBlockRanges(ranges: readonly DirectiveRange[]): BlockPairing[] {
  const orderedBlocks = ranges
    .filter((r) => r.block && (BLOCK_OPENERS.has(r.kind) || BLOCK_CLOSERS.has(r.kind)))
    .slice()
    .sort((a, b) => a.from - b.from);

  const pairings: BlockPairing[] = [];
  const stack: OpenBlock[] = [];
  for (const range of orderedBlocks) {
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
}

function closerHintLabel(kind: BandKind, openerExpr: string): string {
  const head = kind === "each" ? "/each" : "/if";
  return openerExpr ? `${head} · ${openerExpr}` : head;
}

function innermostBlockAt(
  pairings: readonly BlockPairing[],
  caretPos: number | null,
): number | null {
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
}

function computeRailBands(
  pairings: readonly BlockPairing[],
  groups: ProjectedDirectiveGroup[],
  gutter: DirectiveGutterGeometry,
): RailBand[] {
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
}

// ---- Live projection ----------------------------------------------------
const ranges = ref<readonly DirectiveRange[]>([]);
const projectedGroups = ref<ProjectedDirectiveGroup[]>([]);
const gutter = ref<DirectiveGutterGeometry | null>(null);
const caretPos = ref<number | null>(null);
const hoveredBlockId = ref<number | null>(null);

let requestSeq = 0;
let rafId: number | null = null;

function scheduleProject(): void {
  if (rafId !== null) {
    cancelAnimationFrame(rafId);
  }
  rafId = requestAnimationFrame(() => {
    rafId = null;
    void project();
  });
}

async function project(): Promise<void> {
  const seq = ++requestSeq;
  const view = props.getView();
  const pagesContainer = props.getPagesContainer();
  const state = props.editorState;
  if (!view || !pagesContainer || !state) {
    ranges.value = [];
    projectedGroups.value = [];
    gutter.value = null;
    caretPos.value = null;
    return;
  }
  const scanned = getTemplateDirectives(state);
  ranges.value = scanned;
  caretPos.value = scanned.length > 0 ? state.selection.head : null;
  if (scanned.length === 0) {
    projectedGroups.value = [];
    gutter.value = null;
    return;
  }
  gutter.value = measureDirectiveGutter(pagesContainer, props.zoom);
  const projected = await projectRangesToRects<DirectiveRange>(scanned, {
    pagesContainer,
    zoom: props.zoom,
    layout: props.layout,
    blocks: props.blocks,
    measures: props.measures,
  });
  if (seq !== requestSeq) {
    return;
  }
  projectedGroups.value = projected;
}

watch(
  () => [props.editorState, props.zoom, props.layout, props.blocks, props.measures],
  () => scheduleProject(),
);

// ---- Derived pairing / emphasis -----------------------------------------
const pairings = computed(() => pairBlockRanges(ranges.value));
const railBands = computed(() =>
  gutter.value ? computeRailBands(pairings.value, projectedGroups.value, gutter.value) : [],
);
const blockIdByFrom = computed(() => {
  const map = new Map<number, number>();
  for (const pairing of pairings.value) {
    map.set(pairing.openerFrom, pairing.blockId);
    map.set(pairing.closerFrom, pairing.blockId);
  }
  return map;
});
const closerLabelByFrom = computed(() => {
  const map = new Map<number, string>();
  for (const pairing of pairings.value) {
    map.set(pairing.closerFrom, closerHintLabel(pairing.kind, pairing.openerExpr));
  }
  return map;
});
const caretBlockId = computed(() => innermostBlockAt(pairings.value, caretPos.value));

function isActive(blockId: number): boolean {
  return blockId === hoveredBlockId.value || blockId === caretBlockId.value;
}

function isBlockChip(range: DirectiveRange): boolean {
  return range.block && (BLOCK_OPENERS.has(range.kind) || BLOCK_CLOSERS.has(range.kind));
}

function blockIdFor(range: DirectiveRange): number | undefined {
  return range.block ? blockIdByFrom.value.get(range.from) : undefined;
}

function closerTitleFor(range: DirectiveRange): string | undefined {
  return range.block ? closerLabelByFrom.value.get(range.from) : undefined;
}

function chipClass(range: DirectiveRange): (string | Record<string, boolean>)[] {
  const blockId = blockIdFor(range);
  const active = blockId !== undefined && isActive(blockId);
  return [
    "folio-template-directive",
    `folio-template-directive--${range.kind}`,
    { "folio-template-directive--block": isBlockChip(range) },
    { "folio-template-directive--active": active },
  ];
}

// ---- Delegated hover ----------------------------------------------------
function blockIdFromTarget(target: EventTarget | null): number | null {
  if (!(target instanceof HTMLElement)) {
    return null;
  }
  const raw = target.dataset["folioBlockId"];
  return raw === undefined ? null : Number(raw);
}

function handlePointerOver(event: PointerEvent): void {
  const blockId = blockIdFromTarget(event.target);
  if (blockId !== null) {
    hoveredBlockId.value = blockId;
  }
}

function handlePointerOut(event: PointerEvent): void {
  const from = blockIdFromTarget(event.target);
  if (from === null) {
    return;
  }
  const to = blockIdFromTarget(event.relatedTarget);
  if (to === from) {
    return;
  }
  if (hoveredBlockId.value === from) {
    hoveredBlockId.value = to;
  }
}

onMounted(() => scheduleProject());
onBeforeUnmount(() => {
  if (rafId !== null) {
    cancelAnimationFrame(rafId);
    rafId = null;
  }
});
</script>

<style scoped>
.folio-template-directives-overlay {
  position: absolute;
  inset: 0;
  pointer-events: none;
  z-index: 1;
}

/* {{markers}} get Word's "Blue" text-highlight look so fields read like a
   highlighter band over prose. Mirrors React editor.css. */
.folio-template-directive {
  position: absolute;
  border-radius: 3px;
  background: color-mix(in srgb, #0000ff 42%, transparent);
}

.folio-template-directive--clause,
.folio-template-directive--num,
.folio-template-directive--ref {
  background: color-mix(in srgb, #0000ff 24%, transparent);
}

.folio-template-directive--if,
.folio-template-directive--elseif,
.folio-template-directive--else,
.folio-template-directive--endif {
  background: color-mix(in srgb, var(--doc-tmpl-cond, #0d9488) 26%, transparent);
}

.folio-template-directive--each,
.folio-template-directive--endeach {
  background: color-mix(in srgb, var(--doc-tmpl-loop, #7c3aed) 26%, transparent);
}

.folio-template-directive--block {
  pointer-events: auto;
  transition: background-color 120ms ease;
}

.folio-template-directive--if.folio-template-directive--active,
.folio-template-directive--elseif.folio-template-directive--active,
.folio-template-directive--else.folio-template-directive--active,
.folio-template-directive--endif.folio-template-directive--active {
  background: color-mix(in srgb, var(--doc-tmpl-cond, #0d9488) 44%, transparent);
}

.folio-template-directive--each.folio-template-directive--active,
.folio-template-directive--endeach.folio-template-directive--active {
  background: color-mix(in srgb, var(--doc-tmpl-loop, #7c3aed) 44%, transparent);
}

.folio-template-band-rail {
  position: absolute;
  border-radius: 2px;
  pointer-events: auto;
  transition:
    background-color 120ms ease,
    box-shadow 120ms ease;
}

.folio-template-band-rail--each {
  background: color-mix(in srgb, var(--doc-tmpl-loop, #7c3aed) 22%, transparent);
}

.folio-template-band-rail--if {
  background: color-mix(in srgb, var(--doc-tmpl-cond, #0d9488) 22%, transparent);
}

.folio-template-band-rail--each.folio-template-band-rail--active {
  background: color-mix(in srgb, var(--doc-tmpl-loop, #7c3aed) 90%, transparent);
  box-shadow: 0 0 0 0.75px color-mix(in srgb, var(--doc-tmpl-loop, #7c3aed) 45%, transparent);
}

.folio-template-band-rail--if.folio-template-band-rail--active {
  background: color-mix(in srgb, var(--doc-tmpl-cond, #0d9488) 90%, transparent);
  box-shadow: 0 0 0 0.75px color-mix(in srgb, var(--doc-tmpl-cond, #0d9488) 45%, transparent);
}
</style>
