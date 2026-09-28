<template>
  <nav
    v-if="headings.length >= 2"
    ref="navRef"
    :aria-label="outlineLabel"
    class="folio-outline"
    :class="'folio-outline--' + surface"
    :data-folio-outline-surface="surface"
    data-testid="folio-outline"
    :style="{ width: surfaceWidth + 'px' }"
    :tabindex="surface === 'drawer' ? -1 : undefined"
  >
    <button
      v-if="surface === 'rail'"
      type="button"
      class="folio-outline-icon-button"
      data-testid="folio-outline-expand"
      :aria-label="outlineLabel"
      :aria-expanded="expanded"
      :title="outlineLabel"
      @click="emit('expand')"
    >
      <MaterialSymbol name="view_column" :size="16" aria-hidden="true" />
    </button>
    <div v-else class="folio-outline-header">
      <span class="folio-outline-title">{{ t("editor.outlineTitle") }}</span>
      <button
        v-if="surface === 'drawer'"
        type="button"
        class="folio-outline-icon-button"
        :aria-label="t('common.closeDialog')"
        :title="t('common.closeDialog')"
        @click="emit('close')"
      >
        <MaterialSymbol name="close" :size="16" aria-hidden="true" />
      </button>
    </div>
    <OutlineRail
      :items="items"
      :get-scroll-container="getScrollContainer"
      :resolve-pct="resolvePct"
      :on-jump="handleJump"
      :active-id="activeId"
      :presentation="surface === 'rail' ? 'rail' : 'panel'"
      :panel-width="surfaceWidth"
      :top-offset="0"
      :aria-label="outlineLabel"
    />
  </nav>
</template>

<script setup lang="ts">
import { computed, nextTick, ref, watch } from "vue";
import type { HeadingInfo } from "@stll/folio-core/utils/headingCollector";
import { PANEL_METRICS } from "@stll/folio-core/panel-layout";
import { useTranslation } from "../i18n";
import type { OutlineItem } from "../ui/folio-ui";
import { useFolioUI } from "../ui/folio-ui";
import MaterialSymbol from "./ui/MaterialSymbol.vue";

type OutlineSurface = "column" | "rail" | "drawer";

const SURFACE_WIDTH = {
  column: PANEL_METRICS.outlineColumnWidth,
  rail: PANEL_METRICS.outlineRailWidth,
  drawer: PANEL_METRICS.drawerWidth,
} as const satisfies Record<OutlineSurface, number>;

const props = withDefaults(
  defineProps<{
    headings: HeadingInfo[];
    getScrollContainer: () => HTMLElement | null;
    docSize: number;
    activeId: string | null;
    surface: OutlineSurface;
    expanded?: boolean;
  }>(),
  { expanded: false },
);

const emit = defineEmits<{
  navigate: [pmPos: number];
  expand: [];
  close: [];
}>();

const { t } = useTranslation();
const { OutlineRail } = useFolioUI();
const navRef = ref<HTMLElement | null>(null);
const outlineLabel = computed(() => t("editor.showDocumentOutline"));
const surfaceWidth = computed(() => SURFACE_WIDTH[props.surface]);
const items = computed<OutlineItem[]>(() =>
  props.headings.map((heading) => ({
    id: String(heading.pmPos),
    label: heading.text,
    level: heading.level,
    ...(typeof heading.pageNumber === "number" ? { meta: String(heading.pageNumber) } : {}),
  })),
);
const pctById = computed(() => {
  const positions = new Map<string, number>();
  if (props.docSize > 0) {
    for (const heading of props.headings) {
      positions.set(
        String(heading.pmPos),
        Math.min(99, Math.max(1, (heading.pmPos / props.docSize) * 100)),
      );
    }
  }
  return positions;
});

const resolvePct = (id: string) => pctById.value.get(id) ?? null;
const handleJump = (id: string) => {
  emit("navigate", Number(id));
  if (props.surface === "drawer") emit("close");
};

watch(
  () => props.surface === "drawer" && props.headings.length >= 2,
  async (open, _wasOpen, onCleanup) => {
    if (!open || typeof document === "undefined") return;
    const previous = document.activeElement instanceof HTMLElement ? document.activeElement : null;
    let active = true;
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key !== "Escape" || event.defaultPrevented) return;
      event.preventDefault();
      emit("close");
    };
    onCleanup(() => {
      active = false;
      document.removeEventListener("keydown", onKeyDown);
      if (previous?.isConnected) previous.focus({ preventScroll: true });
    });
    await nextTick();
    if (!active) return;
    const nav = navRef.value;
    const target =
      nav?.querySelector<HTMLElement>('[aria-current="true"]') ??
      nav?.querySelector<HTMLElement>('ol button:not([tabindex="-1"])') ??
      nav?.querySelector<HTMLElement>("button") ??
      nav;
    target?.focus({ preventScroll: true });
    document.addEventListener("keydown", onKeyDown);
  },
  { immediate: true, flush: "post" },
);
</script>

<style scoped>
.folio-outline {
  display: flex;
  flex-direction: column;
  box-sizing: border-box;
  min-height: 0;
  background: var(--doc-page, white);
  color: var(--doc-text);
  font-size: 0.8125rem;
}
.folio-outline--column {
  flex: none;
  border-inline-end: 1px solid var(--doc-border);
}
.folio-outline--rail {
  position: relative;
  z-index: 30;
  flex: none;
  align-items: center;
  padding-block: 4px 8px;
  border-inline-end: 1px solid var(--doc-border);
}
.folio-outline--drawer {
  position: absolute;
  inset-block: 0;
  inset-inline-start: 0;
  z-index: 46;
  max-width: calc(100% - 16px);
  border-inline-end: 1px solid var(--doc-border);
  box-shadow: 0 8px 28px var(--doc-shadow-md);
}
.folio-outline-header {
  display: flex;
  flex: none;
  align-items: center;
  justify-content: space-between;
  gap: 0.25rem;
  min-height: 2.25rem;
  padding-block: 0.25rem;
  padding-inline: 0.75rem 0.375rem;
  border-bottom: 1px solid var(--doc-border);
}
.folio-outline-title {
  font-size: 0.75rem;
  font-weight: 600;
  color: var(--doc-text-muted);
}
.folio-outline-icon-button {
  display: inline-flex;
  flex: none;
  align-items: center;
  justify-content: center;
  width: 1.5rem;
  height: 1.5rem;
  padding: 0;
  border: 0;
  border-radius: 0.25rem;
  background: transparent;
  color: var(--doc-text-muted);
  cursor: pointer;
}
.folio-outline-icon-button:hover,
.folio-outline-icon-button[aria-expanded="true"] {
  background: var(--doc-bg-hover);
  color: var(--doc-text);
}
.folio-outline-icon-button:focus-visible {
  outline: 2px solid var(--ring, var(--doc-primary));
  outline-offset: -2px;
}
@media (prefers-reduced-motion: no-preference) {
  .folio-outline--drawer {
    animation: folio-outline-fade-in 0.16s ease-out;
  }
}
@keyframes folio-outline-fade-in {
  from {
    opacity: 0;
  }
  to {
    opacity: 1;
  }
}
</style>
