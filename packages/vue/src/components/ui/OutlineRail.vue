<template>
  <ol
    v-if="items.length >= 2"
    ref="listRef"
    :aria-label="ariaLabel"
    :class="presentation === 'rail' ? 'folio-outline-ticks' : 'folio-outline-list'"
    @keydown="onKeyDown"
    @focusin="focusWithin = true"
    @focusout="onFocusOut"
  >
    <li
      v-for="(item, index) in items"
      :key="item.id"
      :class="presentation === 'rail' ? 'folio-outline-tick-slot' : undefined"
      :style="presentation === 'rail' ? { top: (tickTops[index] ?? 0) + 'px' } : undefined"
    >
      <button
        type="button"
        :class="itemClass(item)"
        :data-depth="Math.min(4, item.level - minLevel)"
        :aria-current="item.id === activeId ? 'true' : undefined"
        :aria-label="presentation === 'rail' ? item.label : undefined"
        :title="presentation === 'panel' ? item.label : undefined"
        :tabindex="index === focusIndex ? 0 : -1"
        @click="jump(item.id)"
        @focus="focusIndex = index"
      >
        <template v-if="presentation === 'rail'">
          <span aria-hidden="true" class="folio-outline-tick-mark" />
          <span aria-hidden="true" class="folio-outline-tick-label">{{ item.label }}</span>
        </template>
        <span v-else class="folio-outline-item-label">{{ item.label }}</span>
      </button>
    </li>
  </ol>
</template>

<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, ref, watch } from "vue";
import type { OutlineItem } from "../../ui/folio-ui";

const TICK_PITCH_PX = 12;

const props = withDefaults(
  defineProps<{
    items: OutlineItem[];
    getScrollContainer: () => HTMLElement | null;
    resolvePct?: (id: string, container: HTMLElement) => number | null;
    onJump: (id: string, container: HTMLElement) => void;
    activeId?: string | null;
    topOffset?: number;
    panelWidth?: number;
    ariaLabel?: string;
    presentation?: "panel" | "rail";
  }>(),
  { activeId: null, ariaLabel: "Outline", presentation: "panel" },
);

const listRef = ref<HTMLOListElement | null>(null);
const height = ref(0);
const focusIndex = ref(0);
const focusWithin = ref(false);
const activeIndex = computed(() => props.items.findIndex((item) => item.id === props.activeId));
const minLevel = computed(() => {
  let minimum = Infinity;
  for (const item of props.items) minimum = Math.min(minimum, item.level);
  return minimum;
});

const itemClass = (item: OutlineItem) => {
  if (props.presentation === "rail") {
    return [
      "folio-outline-tick",
      item.id === props.activeId ? "folio-outline-tick--active" : undefined,
    ];
  }
  return [
    "folio-outline-item",
    item.id === props.activeId ? "folio-outline-item--active" : undefined,
  ];
};

watch(
  activeIndex,
  (index) => {
    if (!focusWithin.value && index >= 0) focusIndex.value = index;
  },
  { immediate: true },
);

watch(
  [activeIndex, () => props.presentation],
  async ([index, presentation]) => {
    if (presentation !== "panel" || index < 0) return;
    await nextTick();
    const list = listRef.value;
    const entry = list?.children.item(index);
    if (!(entry instanceof HTMLElement) || !list) return;
    const top = entry.offsetTop;
    const bottom = top + entry.offsetHeight;
    if (top < list.scrollTop) list.scrollTop = top;
    else if (bottom > list.scrollTop + list.clientHeight)
      list.scrollTop = bottom - list.clientHeight;
  },
  { immediate: true, flush: "post" },
);

let observer: ResizeObserver | null = null;
watch(
  [listRef, () => props.presentation],
  ([list, presentation]) => {
    observer?.disconnect();
    observer = null;
    height.value = 0;
    if (presentation !== "rail" || !list || typeof ResizeObserver === "undefined") return;
    height.value = list.clientHeight;
    observer = new ResizeObserver(() => {
      height.value = list.clientHeight;
    });
    observer.observe(list);
  },
  { flush: "post" },
);
onBeforeUnmount(() => observer?.disconnect());

const tickTops = computed(() => {
  if (props.presentation !== "rail" || height.value === 0) return [];
  const span = Math.max(0, height.value - TICK_PITCH_PX);
  const count = props.items.length;
  if (count * TICK_PITCH_PX > height.value) {
    const step = count > 1 ? span / (count - 1) : 0;
    return props.items.map((_, index) => index * step);
  }
  const container = props.getScrollContainer();
  const tops = props.items.map((item, index) => {
    const pct = container ? (props.resolvePct?.(item.id, container) ?? null) : null;
    return (pct === null ? index / Math.max(1, count - 1) : pct / 100) * span;
  });
  for (let index = 1; index < tops.length; index++) {
    tops[index] = Math.max(tops[index] ?? 0, (tops[index - 1] ?? 0) + TICK_PITCH_PX);
  }
  for (let index = tops.length - 1; index >= 0; index--) {
    const limit = index === tops.length - 1 ? span : (tops[index + 1] ?? span) - TICK_PITCH_PX;
    tops[index] = Math.min(tops[index] ?? 0, limit);
  }
  return tops;
});

const jump = (id: string) => {
  const container = props.getScrollContainer();
  if (container) props.onJump(id, container);
};

const onKeyDown = (event: KeyboardEvent) => {
  const last = props.items.length - 1;
  let next: number;
  switch (event.key) {
    case "ArrowDown":
      next = Math.min(last, focusIndex.value + 1);
      break;
    case "ArrowUp":
      next = Math.max(0, focusIndex.value - 1);
      break;
    case "Home":
      next = 0;
      break;
    case "End":
      next = last;
      break;
    default:
      return;
  }
  event.preventDefault();
  focusIndex.value = next;
  listRef.value?.querySelectorAll<HTMLButtonElement>("button").item(next)?.focus();
};

const onFocusOut = (event: FocusEvent) => {
  if (!(event.relatedTarget instanceof Node && listRef.value?.contains(event.relatedTarget))) {
    focusWithin.value = false;
  }
};
</script>

<style scoped>
.folio-outline-list {
  position: relative;
  flex: 1;
  min-height: 0;
  margin: 0;
  padding: 0.25rem 0;
  overflow-y: auto;
  list-style: none;
}
.folio-outline-item {
  position: relative;
  display: block;
  width: 100%;
  padding-block: 0.3125rem;
  padding-inline: 0.75rem;
  border: 0;
  background: transparent;
  color: var(--doc-text-muted);
  font: inherit;
  text-align: start;
  cursor: pointer;
}
.folio-outline-item[data-depth="1"] {
  padding-inline-start: 1.5rem;
}
.folio-outline-item[data-depth="2"] {
  padding-inline-start: 2.25rem;
}
.folio-outline-item[data-depth="3"] {
  padding-inline-start: 3rem;
}
.folio-outline-item[data-depth="4"] {
  padding-inline-start: 3.75rem;
}
.folio-outline-item-label {
  display: block;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.folio-outline-item:hover,
.folio-outline-item--active {
  background: var(--doc-bg-hover);
  color: var(--doc-text);
}
.folio-outline-item--active {
  font-weight: 600;
}
.folio-outline-item--active::before {
  position: absolute;
  inset-block: 0.1875rem;
  inset-inline-start: 0;
  width: 3px;
  border-radius: 2px;
  background: var(--primary, var(--doc-primary));
  content: "";
}
.folio-outline-ticks {
  position: relative;
  flex: 1;
  width: 100%;
  min-height: 0;
  margin: 0.375rem 0 0;
  padding: 0;
  list-style: none;
}
.folio-outline-tick-slot {
  position: absolute;
  inset-inline: 0;
  height: 12px;
}
.folio-outline-tick {
  position: relative;
  display: flex;
  align-items: center;
  width: 100%;
  height: 100%;
  padding: 0;
  padding-inline-start: 7px;
  border: 0;
  background: transparent;
  cursor: pointer;
}
.folio-outline-tick[data-depth="1"] {
  padding-inline-start: 10px;
}
.folio-outline-tick[data-depth="2"] {
  padding-inline-start: 13px;
}
.folio-outline-tick[data-depth="3"],
.folio-outline-tick[data-depth="4"] {
  padding-inline-start: 16px;
}
.folio-outline-tick-mark {
  display: block;
  width: 14px;
  height: 2px;
  border-radius: 1px;
  background: var(--doc-text-muted);
  opacity: 0.55;
}
.folio-outline-tick[data-depth="1"] .folio-outline-tick-mark {
  width: 11px;
}
.folio-outline-tick[data-depth="2"] .folio-outline-tick-mark,
.folio-outline-tick[data-depth="3"] .folio-outline-tick-mark,
.folio-outline-tick[data-depth="4"] .folio-outline-tick-mark {
  width: 8px;
}
.folio-outline-tick:hover .folio-outline-tick-mark,
.folio-outline-tick:focus-visible .folio-outline-tick-mark {
  opacity: 1;
}
.folio-outline-tick--active .folio-outline-tick-mark {
  height: 3px;
  background: var(--primary, var(--doc-primary));
  opacity: 1;
}
.folio-outline-tick-label {
  position: absolute;
  top: 50%;
  inset-inline-start: calc(100% + 6px);
  display: none;
  max-width: 15rem;
  padding: 0.25rem 0.5rem;
  overflow: hidden;
  border: 1px solid var(--doc-border);
  border-radius: 0.375rem;
  background: var(--popover, var(--doc-page));
  box-shadow: 0 4px 12px var(--doc-shadow-sm);
  color: var(--popover-foreground, var(--doc-text));
  font-size: 0.75rem;
  line-height: 1.25;
  text-overflow: ellipsis;
  white-space: nowrap;
  transform: translateY(-50%);
  pointer-events: none;
}
.folio-outline-tick:hover .folio-outline-tick-label,
.folio-outline-tick:focus-visible .folio-outline-tick-label {
  display: block;
}
.folio-outline-item:focus-visible,
.folio-outline-tick:focus-visible {
  outline: 2px solid var(--ring, var(--doc-primary));
  outline-offset: -2px;
}
</style>
