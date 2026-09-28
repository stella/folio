<!--
  UnifiedSidebar — anchored cards next to the rendered document.
  Mirrors React's UnifiedSidebar.tsx visual model: cards live in a
  fixed-width column to the right of the page, each card sits at the
  Y of its corresponding [data-comment-id] / .docx-insertion span.
  Falls back to stacked layout when an anchor can't be resolved (e.g.
  the layout-painter hasn't finished rendering yet) — cards still
  show up rather than going invisible.

  The single source of truth for the item list is
  `useCommentSidebarItems({ comments, trackedChanges, ... })`, fed from this
  component's props; the template dispatches each item to the matching card.
-->
<template>
  <!-- Dynamic style boost for the focused/expanded item — same
       approach React's DocxEditor.tsx:5029-5044 takes. Injected as
       a sibling so the !important overrides win against the base
       editor.css highlight rules without touching DOM nodes. -->
  <component v-if="expandedHighlightCss" :is="'style'">{{ expandedHighlightCss }}</component>
  <aside
    v-if="isOpen"
    ref="rootRef"
    :class="['unified-sidebar', { 'unified-sidebar--drawer': isDrawer }]"
    :data-folio-comments-surface="surface ?? 'column'"
    data-testid="folio-comments"
    :tabindex="isDrawer ? -1 : undefined"
    :style="asideStyle"
    @mousedown="onSidebarMouseDown"
  >
    <div
      class="unified-sidebar__inner"
      :style="isDrawer ? drawerInnerStyle : { minHeight: minHeightPx + 'px' }"
    >
      <!-- Every item — add-comment input, comments, tracked changes —
           flows through the same `items` list and the shared
           `resolveItemPositions` collision pass (mirrors React's
           UnifiedSidebar.tsx). The add-comment card no longer has a
           separate, independently-positioned block, so it claims its Y
           slot and neighbouring cards stack below it instead of
           overlapping (fixes #669). -->
      <template v-for="item in items" :key="item.id">
        <div
          class="unified-sidebar__card-slot"
          :data-card-id="item.id"
          :style="cardSlotStyle(item.id)"
        >
          <AddCommentCard
            v-if="item.kind === 'add-comment'"
            @submit="(text: string) => $emit('add-comment', text)"
            @cancel="$emit('cancel-add-comment')"
          />
          <!-- Resolved + collapsed comments render as a small
               chat-bubble-check marker (matches React's
               useCommentSidebarItems.tsx:96-98). Click expands into
               the full card. -->
          <ResolvedCommentMarker
            v-else-if="item.kind === 'comment' && item.comment!.done && expandedId !== item.id"
            :comment="item.comment!"
            @toggle-expand="toggleExpanded(item.id)"
          />
          <CommentCard
            v-else-if="item.kind === 'comment'"
            :comment="item.comment!"
            :replies="item.replies!"
            :expanded="expandedId === item.id"
            @click="toggleExpanded(item.id)"
            @reply="(id: number, text: string) => $emit('comment-reply', id, text)"
            @resolve="(id: number) => $emit('comment-resolve', id)"
            @unresolve="(id: number) => $emit('comment-unresolve', id)"
            @delete="(id: number) => $emit('comment-delete', id)"
          />
          <TrackedChangeCard
            v-else-if="item.kind === 'tracked-change'"
            :change="item.change!"
            :replies="item.replies ?? []"
            :expanded="expandedId === item.id"
            @click="toggleExpanded(item.id)"
            @accept="(from: number, to: number) => $emit('accept-change', from, to)"
            @reject="(from: number, to: number) => $emit('reject-change', from, to)"
            @accept-by-id="(rev: number) => $emit('accept-change-by-id', rev)"
            @reject-by-id="(rev: number) => $emit('reject-change-by-id', rev)"
            @reply="(rev: number, text: string) => $emit('tracked-change-reply', rev, text)"
          />
        </div>
      </template>
    </div>
  </aside>
</template>

<script setup lang="ts">
import { ref, computed, toRef, watch, onMounted, onBeforeUnmount, type CSSProperties } from "vue";
import {
  commentAnchorSelector,
  indexCommentAnchors,
} from "@stll/folio-core/render-dom/commentAnchorAttributes";
import type { Comment } from "@stll/folio-core/types/content";
import type { TrackedChangeEntry } from "./sidebar/sidebarUtils";
import { createRenderedDomContext } from "@stll/folio-core/render-dom/RenderedDomContext";
import { resolveSidebarItemPositions } from "@stll/folio-core/render-dom/resolveSidebarItemPositions";
import { PANEL_METRICS } from "@stll/folio-core/panel-layout";
import CommentCard from "./sidebar/CommentCard.vue";
import ResolvedCommentMarker from "./sidebar/ResolvedCommentMarker.vue";
import TrackedChangeCard from "./sidebar/TrackedChangeCard.vue";
import AddCommentCard from "./sidebar/AddCommentCard.vue";
import { useCommentSidebarItems } from "../composables/useCommentSidebarItems";

const props = defineProps<{
  isOpen: boolean;
  surface?: "column" | "drawer";
  comments: Comment[];
  trackedChanges: TrackedChangeEntry[];
  isAddingComment?: boolean;
  showResolved?: boolean;
  pagesContainer: HTMLElement | null;
  pageWidthPx: number;
  zoom?: number;
  /** Controlled expand: when set, overrides local click toggling.
   *  Used by DocxEditor to auto-expand cards when the cursor
   *  lands on a commented / tracked span (mirrors React
   *  DocxEditor.tsx:5080-5118 cursorSidebarItem detection). */
  activeItemId?: string | null;
  /** Y (in unscaled coords inside the pages-viewport) where the
   *  AddCommentCard should anchor — mirrors React's
   *  addCommentYPosition pass-through. Null = top of rail. */
  addCommentYPosition?: number | null;
}>();

const emit = defineEmits<{
  (e: "close"): void;
  (e: "dismiss"): void;
  (e: "add-comment", text: string): void;
  (e: "cancel-add-comment"): void;
  (e: "comment-reply", commentId: number, text: string): void;
  (e: "comment-resolve", commentId: number): void;
  (e: "comment-unresolve", commentId: number): void;
  (e: "comment-delete", commentId: number): void;
  (e: "accept-change", from: number, to: number): void;
  (e: "reject-change", from: number, to: number): void;
  /** For paragraph-mark and other structural revisions — accept/reject by w:id. */
  (e: "accept-change-by-id", revisionId: number): void;
  (e: "reject-change-by-id", revisionId: number): void;
  (e: "tracked-change-reply", revisionId: number, text: string): void;
  (e: "update:activeItemId", id: string | null): void;
}>();

const isDrawer = computed(() => props.surface === "drawer");

// Local fallback for uncontrolled use; when `activeItemId` is bound
// from the parent (DocxEditor) the prop wins and toggleExpanded
// emits up so the parent can update its own state.
const localExpanded = ref<string | null>(null);
const expandedId = computed<string | null>(() =>
  props.activeItemId !== undefined ? props.activeItemId : localExpanded.value,
);

function toggleExpanded(id: string) {
  const next = expandedId.value === id ? null : id;
  localExpanded.value = next;
  emit("update:activeItemId", next);
}

// Always stop sidebar mousedowns from reaching the editor (was `@mousedown.stop`,
// which prevents the click from moving the PM cursor / stealing focus). On top
// of that, clicking the empty sidebar background — anywhere that isn't a card
// slot — collapses the expanded item, matching React (where clicking the grey
// gutter behind the cards deselects). Card clicks are handled by each card.
function onSidebarMouseDown(e: MouseEvent) {
  e.stopPropagation();
  const target = e.target;
  if (target instanceof HTMLElement && target.closest(".unified-sidebar__card-slot")) return;
  if (expandedId.value !== null) {
    localExpanded.value = null;
    emit("update:activeItemId", null);
  }
}

// Single source of truth: derive the flat sidebar item list from comments +
// tracked changes. `resolveItemPositions` and the template dispatch below
// consume the same list.
const items = useCommentSidebarItems({
  comments: toRef(props, "comments"),
  trackedChanges: toRef(props, "trackedChanges"),
  showResolved: computed(() => props.showResolved ?? false),
  isAddingComment: computed(() => props.isAddingComment ?? false),
  addCommentYPosition: computed(() => props.addCommentYPosition ?? null),
});

// Resolved Y per item id. Recomputed on tick changes (manual recompute,
// ResizeObserver firing, watch on items length). Falls back to stacked
// layout when an anchor isn't found yet.
const rootRef = ref<HTMLElement | null>(null);
const resolvedY = ref<Map<string, number>>(new Map());
// Persistent across recomputes: lets resolveItemPositions keep a card
// at its last-known Y during transient layout instead of popping it out.
const lastKnown = new Map<string, number>();
let resizeObserver: ResizeObserver | null = null;
// Observes every card slot. A card grows when it expands (reply input +
// thread mount) or when its reply textarea auto-grows; the pagesContainer
// observer never sees that, so without this the cards below stay stacked at
// the collapsed height and the expanded card overlaps its neighbour.
// Observing the slots re-runs the collision pass on any height change.
let cardResizeObserver: ResizeObserver | null = null;
// The slot elements currently observed — keyed by element identity, NOT by
// card id. After a sidebar close/reopen the same ids reappear on brand-new
// DOM nodes, so an id-string guard would keep observing detached nodes;
// comparing elements re-binds to the live ones. Re-`observe()` is skipped
// when the element set is unchanged (it would otherwise re-fire the
// initial callback and spin recompute).
let observedSlots = new Set<HTMLElement>();

function syncCardObservers() {
  const root = rootRef.value;
  if (!root || !cardResizeObserver) return;
  const slots = new Set(root.querySelectorAll<HTMLElement>("[data-card-id]"));
  if (slots.size === observedSlots.size && [...slots].every((el) => observedSlots.has(el))) {
    return;
  }
  cardResizeObserver.disconnect();
  for (const el of slots) cardResizeObserver.observe(el);
  observedSlots = slots;
}

function computePositions() {
  updatePanelGeometry();
  if (isDrawer.value) {
    resolvedY.value = new Map();
    return;
  }
  const container = props.pagesContainer;
  const list = items.value;
  if (!container || list.length === 0) {
    resolvedY.value = new Map();
    return;
  }

  // ONE batched DOM read: build maps for comments/insertions/deletions
  // up front, then look each item up by id rather than running N
  // querySelectors per recompute.
  const containerRect = container.getBoundingClientRect();
  // A run inside overlapping ranges is indexed under every comment it is in.
  const commentEls = indexCommentAnchors(container);
  const insertionEls = new Map<string, HTMLElement>();
  for (const el of container.querySelectorAll<HTMLElement>(".docx-insertion[data-revision-id]")) {
    const id = el.dataset["revisionId"];
    if (id && !insertionEls.has(id)) insertionEls.set(id, el);
  }
  const deletionEls = new Map<string, HTMLElement>();
  for (const el of container.querySelectorAll<HTMLElement>(".docx-deletion[data-revision-id]")) {
    const id = el.dataset["revisionId"];
    if (id && !deletionEls.has(id)) deletionEls.set(id, el);
  }
  // Structural tracked changes (whole-table / row / cell insert+delete +
  // tracked paragraph marks) live on the painted table/row/cell or on
  // paragraph-fragment elements, not on `.docx-insertion` text spans —
  // without these, an empty inserted table, a cell-only insert, or a
  // pure-pmark revision never anchors and its card stays invisible.
  // Two class prefixes are in play: `ep-revision-*` for table scopes,
  // `layout-revision-*` for paragraph marks (renderParagraph.ts:128).
  for (const el of container.querySelectorAll<HTMLElement>(
    ".ep-revision-table[data-revision-id], " +
      ".ep-revision-row[data-revision-id], " +
      ".ep-revision-cell[data-revision-id], " +
      ".layout-revision-pmark[data-revision-id]",
  )) {
    const id = el.dataset["revisionId"];
    if (!id) continue;
    const isIns =
      el.classList.contains("ep-revision-ins") || el.classList.contains("layout-revision-ins");
    const map = isIns ? insertionEls : deletionEls;
    if (!map.has(id)) map.set(id, el);
  }

  // Resolve each anchored item's Y from its painted span and key it by
  // the item's anchorKey (`comment-<id>` / `revision-<revId>`), which is
  // what resolveItemPositions looks up. The add-comment item carries a
  // fixedY instead and needs no DOM anchor. Y is in pages-container
  // coords, already post-zoom (getBoundingClientRect is post-transform),
  // so resolveItemPositions runs with zoom 1.
  const anchorPositions = new Map<string, number>();
  const anchorY = (el: HTMLElement) =>
    el.getBoundingClientRect().top - containerRect.top + container.scrollTop;
  for (const item of list) {
    if (!item.anchorKey) continue;
    let anchor: HTMLElement | undefined;
    if (item.kind === "comment") {
      anchor = commentEls.get(String(item.comment!.id));
    } else if (item.kind === "tracked-change") {
      const change = item.change!;
      anchor =
        change.type === "deletion"
          ? deletionEls.get(String(change.revisionId))
          : insertionEls.get(String(change.insertionRevisionId ?? change.revisionId));
    }
    if (anchor) anchorPositions.set(item.anchorKey, anchorY(anchor));
  }

  // Card-height lookup: also batched into one querySelectorAll.
  const cardHeights = new Map<string, number>();
  const root = rootRef.value;
  if (root) {
    for (const el of root.querySelectorAll<HTMLElement>("[data-card-id]")) {
      const id = el.dataset["cardId"];
      if (id) cardHeights.set(id, el.offsetHeight);
    }
  }

  const map = new Map<string, number>();
  for (const { item, y } of resolveSidebarItemPositions({
    items: list,
    anchorPositions,
    renderedDomContext: createRenderedDomContext(container),
    zoom: 1,
    cardHeights,
    lastKnown,
  })) {
    map.set(item.id, y);
  }
  resolvedY.value = map;

  // Cards are in the DOM now — observe each slot so a later height change
  // (expand, reply thread render, textarea growth) re-runs this pass.
  syncCardObservers();
}

const minHeightPx = computed(() => {
  let max = 0;
  for (const y of resolvedY.value.values()) max = Math.max(max, y);
  return max + 200; // headroom for the bottom card
});

// The page and sidebar share the pages viewport. Measure the painted page
// edge so a reserved comments track, zoom, and horizontal scrolling all use
// the same position that the user sees.
const measuredLeft = ref<number | null>(null);
const initialDrawerBox: { top: number; width: number; height: number } = {
  top: 0,
  width: PANEL_METRICS.drawerWidth,
  height: 0,
};
const drawerBox = ref(initialDrawerBox);

function updatePanelGeometry() {
  const root = rootRef.value;
  const parent = root?.offsetParent;
  const scroll = findScrollParent(props.pagesContainer);
  if (!(parent instanceof HTMLElement) || !scroll) return;

  const parentRect = parent.getBoundingClientRect();
  const scrollRect = scroll.getBoundingClientRect();
  if (isDrawer.value) {
    const width = Math.max(0, Math.min(PANEL_METRICS.drawerWidth, scroll.clientWidth - 16));
    drawerBox.value = {
      top: scrollRect.top - parentRect.top,
      width,
      height: scroll.clientHeight,
    };
    measuredLeft.value = scrollRect.left + scroll.clientWidth - parentRect.left - width;
    return;
  }

  const page = props.pagesContainer?.querySelector<HTMLElement>(".layout-page");
  if (!page) {
    measuredLeft.value = null;
    return;
  }
  const rawLeft = page.getBoundingClientRect().right - parentRect.left + PANEL_METRICS.commentsGap;
  const maxVisibleLeft = Math.max(8, parent.clientWidth - PANEL_METRICS.commentsWidth - 8);
  measuredLeft.value = Math.max(8, Math.min(rawLeft, maxVisibleLeft));
}

const drawerInnerStyle = computed<CSSProperties>(() => ({
  position: "sticky",
  top: 0,
  maxHeight: drawerBox.value.height || undefined,
  overflowY: "auto",
  paddingBlock: PANEL_METRICS.commentsGap + "px",
  boxSizing: "border-box",
}));

// Dynamic CSS boost for the expanded item. Mirrors React
// DocxEditor.tsx:5029-5044: brighten the comment-anchor highlight
// (yellow) for the focused comment, and the tracked-change
// insertion/deletion spans for the focused tc card.
const expandedHighlightCss = computed(() => {
  const id = expandedId.value;
  if (!id) return "";
  if (id.startsWith("comment-")) {
    const cid = id.slice("comment-".length);
    return `${commentAnchorSelector(cid, ".paged-editor__pages ")} { background-color: rgba(255, 212, 0, 0.35) !important; border-bottom: 2px solid rgba(255, 212, 0, 0.7) !important; }`;
  }
  if (id.startsWith("tc-")) {
    // id shape: tc-<revisionId>-<index>
    const parts = id.split("-");
    const revId = parts.at(1) ?? "";
    const item = items.value.find((s) => s.id === id);
    const insRev = item?.change?.insertionRevisionId ?? Number(revId);
    return `
      .paged-editor__pages .docx-insertion[data-revision-id="${insRev}"] { background-color: rgba(52, 168, 83, 0.2) !important; border-bottom: 2px solid #2e7d32 !important; }
      .paged-editor__pages .docx-deletion[data-revision-id="${revId}"] { background-color: rgba(211, 47, 47, 0.2) !important; text-decoration-thickness: 2px !important; }
    `;
  }
  return "";
});

const asideStyle = computed<CSSProperties>(() => {
  const drawer = isDrawer.value;
  const hasPositions = drawer
    ? measuredLeft.value !== null
    : resolvedY.value.size > 0 || items.value.length === 0;
  return {
    position: "absolute",
    top: drawer ? drawerBox.value.top + "px" : "0",
    left:
      measuredLeft.value === null
        ? `calc(50% + ${props.pageWidthPx / 2 + PANEL_METRICS.commentsGap}px)`
        : measuredLeft.value + "px",
    width: (drawer ? drawerBox.value.width : PANEL_METRICS.commentsWidth) + "px",
    paddingInline: drawer ? PANEL_METRICS.commentsGap + "px" : undefined,
    bottom: drawer ? undefined : 0,
    opacity: hasPositions ? 1 : 0,
    pointerEvents: hasPositions ? "auto" : "none",
  };
});

function cardSlotStyle(id: string): CSSProperties {
  if (isDrawer.value) return { position: "relative", marginBottom: "8px" };
  const y = resolvedY.value.get(id);
  if (y == null) {
    // Fall back to stacked layout: card flows naturally below the
    // previous one, fully visible (NOT opacity 0). The user always
    // sees comments even when anchor measurement hasn't settled.
    return {
      position: "static",
      marginBottom: "8px",
    };
  }
  return {
    position: "absolute",
    top: y + "px",
    left: 0,
    right: 0,
    transition: "top 0.15s ease",
  };
}

// Schedule a single recompute on the next animation frame, coalesced.
let scheduled = false;
function recompute() {
  if (scheduled) return;
  scheduled = true;
  requestAnimationFrame(() => {
    scheduled = false;
    computePositions();
  });
}

// Items / expand state / container / zoom changes all bump positions
// on the next animation frame. Zoom matters because the page is
// `transform: scale(zoom)` and `getBoundingClientRect` returns
// post-transform coords, so a zoom change shifts every anchor.
watch(
  () => [
    items.value,
    expandedId.value,
    props.pagesContainer,
    props.pageWidthPx,
    props.zoom,
    props.isAddingComment,
    props.addCommentYPosition,
    props.surface,
    props.isOpen,
  ],
  () => recompute(),
  { immediate: true },
);

// Find the closest scrolling ancestor of pagesContainer — usually the
// pages-viewport — so we can re-run computePositions whenever the user
// scrolls. Without this listener cards stay at their stale absolute Y
// while anchors move with the scrolled content; comments visibly drift
// out of sync as the user scrolls.
function findScrollParent(el: HTMLElement | null): HTMLElement | null {
  let cur: HTMLElement | null = el?.parentElement ?? null;
  while (cur) {
    const overflowY = getComputedStyle(cur).overflowY;
    if (overflowY === "auto" || overflowY === "scroll") return cur;
    cur = cur.parentElement;
  }
  return null;
}

let scrollParent: HTMLElement | null = null;
function bindScrollListener() {
  if (scrollParent) scrollParent.removeEventListener("scroll", recompute);
  scrollParent = findScrollParent(props.pagesContainer);
  if (scrollParent) {
    scrollParent.addEventListener("scroll", recompute, { passive: true });
    resizeObserver?.observe(scrollParent);
  }
}

onMounted(() => {
  // Watches card-slot height changes (expand / reply thread / textarea).
  // computePositions() binds the observations once cards render.
  cardResizeObserver = new ResizeObserver(() => recompute());
  recompute();
  // Bind ResizeObserver once pagesContainer is non-null.
  if (props.pagesContainer) {
    resizeObserver = new ResizeObserver(() => recompute());
    resizeObserver.observe(props.pagesContainer);
    bindScrollListener();
  }
  window.addEventListener("resize", recompute);
});

watch(
  () => props.pagesContainer,
  (el) => {
    resizeObserver?.disconnect();
    resizeObserver = null;
    if (el) {
      resizeObserver = new ResizeObserver(() => recompute());
      resizeObserver.observe(el);
    }
    bindScrollListener();
    recompute();
  },
);

onBeforeUnmount(() => {
  resizeObserver?.disconnect();
  cardResizeObserver?.disconnect();
  if (scrollParent) scrollParent.removeEventListener("scroll", recompute);
  window.removeEventListener("resize", recompute);
});

watch(
  () => props.isOpen && isDrawer.value,
  (active, _previous, onCleanup) => {
    if (!active || typeof document === "undefined") return;
    const previousFocus =
      document.activeElement instanceof HTMLElement ? document.activeElement : null;
    rootRef.value?.focus({ preventScroll: true });
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key !== "Escape" || event.defaultPrevented) return;
      event.preventDefault();
      emit("dismiss");
    };
    document.addEventListener("keydown", onKeyDown);
    onCleanup(() => {
      document.removeEventListener("keydown", onKeyDown);
      if (previousFocus?.isConnected) previousFocus.focus({ preventScroll: true });
    });
  },
  { immediate: true, flush: "post" },
);
</script>

<style scoped>
.unified-sidebar {
  background: transparent;
  font-family: "Google Sans", Roboto, Arial, sans-serif;
  pointer-events: auto;
  z-index: 5;
  transition: opacity 0.15s ease;
}
.unified-sidebar--drawer {
  box-sizing: border-box;
  background: var(--doc-canvas-surface, var(--doc-page, white));
  border-inline-start: 1px solid var(--doc-border);
  box-shadow: 0 8px 28px var(--doc-shadow-md);
  outline: none;
  z-index: 46;
}
.unified-sidebar__inner {
  position: relative;
}
.unified-sidebar:not(.unified-sidebar--drawer) .unified-sidebar__inner {
  padding-inline: 0;
}
</style>
