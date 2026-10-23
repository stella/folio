/**
 * Dropdown-menu state for the Vue Toolbar — open/closed state, position
 * tracking via `position: fixed` (the `.basic-toolbar` parent forces
 * `overflow-x: auto` which clips `position: absolute` children), and
 * outside-click dismissal.
 *
 * Mirrors the house pattern from `useFixedDropdown.ts` but tailored to a
 * shared `openDropdown` ref that gates which named menu is visible.
 */

import {
  computed,
  type CSSProperties,
  onBeforeUnmount,
  onMounted,
  type Ref,
  ref,
} from "vue";

import { onDocumentPointerDown } from "@stll/folio-core/utils/documentPointerDown";

export type UseToolbarDropdownsOptions = {
  zoom: Ref<HTMLElement | null>;
  style: Ref<HTMLElement | null>;
  font: Ref<HTMLElement | null>;
  size: Ref<HTMLElement | null>;
  align: Ref<HTMLElement | null>;
  spacing: Ref<HTMLElement | null>;
};

export function useToolbarDropdowns(refs: UseToolbarDropdownsOptions) {
  const openDropdown = ref<string | null>(null);
  const dropdownPos = ref<{ top: number; left: number }>({ top: 0, left: 0 });

  function dropdownTriggerRef(name: string): HTMLElement | null {
    switch (name) {
      case "zoom":
        return refs.zoom.value;
      case "style":
        return refs.style.value;
      case "font":
        return refs.font.value;
      case "size":
        return refs.size.value;
      case "align":
        return refs.align.value;
      case "spacing":
        return refs.spacing.value;
      default:
        return null;
    }
  }

  function recomputeDropdownPos(name: string) {
    const el = dropdownTriggerRef(name);
    if (!el) return;
    const r = el.getBoundingClientRect();
    dropdownPos.value = { top: r.bottom, left: r.left };
  }

  const dropdownMenuStyle = computed<CSSProperties>(() => ({
    top: `${dropdownPos.value.top}px`,
    left: `${dropdownPos.value.left}px`,
  }));

  function toggleDropdown(name: string) {
    if (openDropdown.value !== name) recomputeDropdownPos(name);
    openDropdown.value = openDropdown.value === name ? null : name;
  }

  function closeDropdowns(e: PointerEvent) {
    const target = e.target;
    if (!(target instanceof Node)) return;
    const allRefs = [
      refs.zoom,
      refs.style,
      refs.font,
      refs.size,
      refs.align,
      refs.spacing,
    ];
    if (!allRefs.some((r) => r.value?.contains(target))) {
      openDropdown.value = null;
    }
  }

  let detachPointerDown: (() => void) | undefined;
  onMounted(() => {
    detachPointerDown = onDocumentPointerDown(document, closeDropdowns);
  });
  onBeforeUnmount(() => detachPointerDown?.());

  return {
    openDropdown,
    dropdownMenuStyle,
    toggleDropdown,
    recomputeDropdownPos,
  };
}
