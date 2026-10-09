import { computed, nextTick, toValue, watch } from "vue";
import type { MaybeRefOrGetter, Ref } from "vue";

export type OutlineSurface = "column" | "rail" | "drawer";

type UseOutlineDrawerFocusOptions = {
  surface: MaybeRefOrGetter<OutlineSurface>;
  available: MaybeRefOrGetter<boolean>;
  navRef: Ref<HTMLElement | null>;
  onClose: () => void;
};

/** Focus and Escape behavior follows the same availability gate as the panel DOM. */
export const useOutlineDrawerFocus = ({
  surface,
  available,
  navRef,
  onClose,
}: UseOutlineDrawerFocusOptions): void => {
  const isOpen = computed(() => toValue(surface) === "drawer" && toValue(available));

  watch(
    isOpen,
    async (open, _wasOpen, onCleanup) => {
      if (!open || typeof document === "undefined") return;
      const previous =
        document.activeElement instanceof HTMLElement ? document.activeElement : null;
      let active = true;
      const onKeyDown = (event: KeyboardEvent) => {
        if (event.key !== "Escape" || event.defaultPrevented) return;
        event.preventDefault();
        onClose();
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
};
