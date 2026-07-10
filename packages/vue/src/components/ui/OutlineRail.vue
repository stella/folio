<!--
  Built-in, dependency-light OutlineRail used when a consumer does not inject
  one. Vue port of packages/react/src/ui/defaults/outline-rail.tsx: a plain
  list of clickable entries indented by level; clicking an entry calls `onJump`
  with the resolved scroll container.

  Renders a bare `<ul>` (not a `<nav>` wrapper, unlike React's default): the
  one Vue consumer (DocumentOutline.vue) already renders its own `<nav>`
  landmark (and its own empty-state message) around this list, so a default
  `<nav>` or a "hide below two items" gate here would duplicate/override that
  wrapper's behavior. Unlike React's `DefaultOutlineRail` (which no-ops below
  two items), this default renders whenever there is at least one item.
  `ariaLabel` is applied to the list itself.
-->
<template>
  <ul v-if="items.length > 0" class="folio-default-outline-list" :aria-label="ariaLabel">
    <li v-for="item in items" :key="item.id">
      <button
        type="button"
        class="folio-default-outline-item"
        :class="{ 'folio-default-outline-item--active': item.id === activeId }"
        :aria-current="item.id === activeId ? 'true' : undefined"
        :style="{ paddingInlineStart: `${item.level * 16 + 8}px` }"
        :title="item.label"
        @mousedown.prevent="onClick(item.id)"
      >
        {{ item.label }}
      </button>
    </li>
  </ul>
</template>

<script setup lang="ts">
import type { OutlineItem } from "../../ui/folio-ui";

const props = defineProps<{
  items: OutlineItem[];
  getScrollContainer: () => HTMLElement | null;
  onJump: (id: string, container: HTMLElement) => void;
  activeId?: string | null;
  ariaLabel?: string;
}>();

function onClick(id: string) {
  const container = props.getScrollContainer();
  if (container) props.onJump(id, container);
}
</script>

<style scoped>
.folio-default-outline-list {
  list-style: none;
  margin: 0;
  padding: 0;
}
.folio-default-outline-item {
  display: block;
  width: 100%;
  text-align: left;
  padding-block: 6px;
  padding-inline-end: 12px;
  border: none;
  background: transparent;
  cursor: pointer;
  font-size: 13px;
  color: var(--doc-text-muted, inherit);
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
  border-radius: 4px;
}
.folio-default-outline-item:hover {
  background: var(--doc-shadow-subtle, rgb(0 0 0 / 6%));
}
.folio-default-outline-item--active {
  color: var(--doc-primary, #2563eb);
  font-weight: 600;
}
</style>
