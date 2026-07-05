/**
 * Vue port of packages/react/src/hooks/useSelectionHighlight.ts.
 * Wraps the framework-agnostic selection helpers (ported into the Vue
 * adapter's `utils/selectionHighlight`). Same debounce semantics (16ms
 * default), same overlay rect calculation, same selection-within-container
 * detection.
 *
 * NOTE: upstream imported a no-arg DOM `getSelectedText()` from core's
 * `utils`; our core only exposes a ProseMirror-state `getSelectedText(state)`,
 * so the DOM read is inlined here (`window.getSelection().toString()`).
 */
import {
  ref,
  computed,
  watch,
  onMounted,
  onBeforeUnmount,
  type Ref,
  type ComputedRef,
  type CSSProperties,
} from "vue";
import {
  type HighlightRect,
  type SelectionHighlightConfig,
  DEFAULT_SELECTION_STYLE,
  getMergedSelectionRects,
  hasActiveSelection,
  isSelectionWithin,
  injectSelectionStyles,
  areSelectionStylesInjected,
} from "../utils/selectionHighlight";

/** DOM-selection text read (replaces core's no-arg upstream helper). */
function getSelectedText(): string {
  return window.getSelection()?.toString() ?? "";
}

export type UseSelectionHighlightOptions = {
  containerRef: Ref<HTMLElement | null>;
  enabled?: boolean;
  config?: SelectionHighlightConfig;
  useOverlay?: boolean;
  debounceMs?: number;
  onSelectionChange?: (hasSelection: boolean, text: string) => void;
}

export type UseSelectionHighlightReturn = {
  hasSelection: ComputedRef<boolean>;
  selectedText: ComputedRef<string>;
  highlightRects: ComputedRef<HighlightRect[]>;
  isSelectionInContainer: ComputedRef<boolean>;
  refresh: () => void;
  getOverlayStyle: (rect: HighlightRect) => CSSProperties;
}

export function useSelectionHighlight(
  options: UseSelectionHighlightOptions,
): UseSelectionHighlightReturn {
  const {
    containerRef,
    enabled = true,
    config = DEFAULT_SELECTION_STYLE,
    useOverlay = false,
    debounceMs = 16,
    onSelectionChange,
  } = options;

  const hasSelectionState = ref(false);
  const selectedText = ref("");
  const highlightRects = ref<HighlightRect[]>([]);
  const isSelectionInContainer = ref(false);

  let debounceTimer: ReturnType<typeof setTimeout> | null = null;
  let lastUpdate = 0;

  function update() {
    const container = containerRef.value;
    const active = hasActiveSelection();
    const text = getSelectedText();
    const inContainer = container ? isSelectionWithin(container) : false;
    hasSelectionState.value = active;
    selectedText.value = text;
    isSelectionInContainer.value = inContainer;
    highlightRects.value =
      useOverlay && inContainer && container ? getMergedSelectionRects(container) : [];
    onSelectionChange?.(active && inContainer, text);
  }

  function refresh() {
    update();
  }

  function debouncedUpdate() {
    const now = performance.now();
    if (now - lastUpdate < debounceMs) {
      if (debounceTimer) clearTimeout(debounceTimer);
      debounceTimer = setTimeout(() => {
        lastUpdate = performance.now();
        update();
        debounceTimer = null;
      }, debounceMs);
      return;
    }
    lastUpdate = now;
    update();
  }

  function getOverlayStyle(rect: HighlightRect): CSSProperties {
    const style: CSSProperties = {
      position: "absolute",
      left: `${rect.left}px`,
      top: `${rect.top}px`,
      width: `${rect.width}px`,
      height: `${rect.height}px`,
      backgroundColor: config.backgroundColor,
      zIndex: config.zIndex ?? 0,
      pointerEvents: "none",
    };
    if (config.borderRadius) style.borderRadius = `${config.borderRadius}px`;
    if (config.borderColor) style.border = `1px solid ${config.borderColor}`;
    return style;
  }

  onMounted(() => {
    if (!enabled) return;
    if (!areSelectionStylesInjected()) injectSelectionStyles(config);
    document.addEventListener("selectionchange", debouncedUpdate);
  });

  onBeforeUnmount(() => {
    document.removeEventListener("selectionchange", debouncedUpdate);
    if (debounceTimer) clearTimeout(debounceTimer);
  });

  watch(containerRef, () => update());

  return {
    hasSelection: computed(() => hasSelectionState.value),
    selectedText: computed(() => selectedText.value),
    highlightRects: computed(() => highlightRects.value),
    isSelectionInContainer: computed(() => isSelectionInContainer.value),
    refresh,
    getOverlayStyle,
  };
}
