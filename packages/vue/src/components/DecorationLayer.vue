<!--
  PORT-BLOCKED: `../plugin-api/RenderedDomContext` (createRenderedDomContext) is
  absent from folio — the whole plugin-api subsystem is not yet ported to
  @stll/folio-core (see CORE-API-MAP "UPSTREAM-ONLY-FEATURE"). It is stubbed
  locally below to a no-op context so this overlay compiles and mounts without
  painting decorations. TODO: replace the stub with the real RenderedDomContext
  once core's plugin-api lands.
-->
<template>
  <div ref="overlayRef" class="paged-editor__decoration-overlay" aria-hidden="true" />
</template>

<script setup lang="ts">
import { onBeforeUnmount, onMounted, ref, watch } from "vue";
import type { EditorState } from "prosemirror-state";
import type { Decoration, DecorationSet, EditorView } from "prosemirror-view";
import type { LayoutSelectionGate } from "@stll/folio-core/paged-layout/LayoutSelectionGate";

const props = defineProps<{
  getView: () => EditorView | null;
  getPagesContainer: () => HTMLElement | null;
  zoom: number;
  transactionVersion: number;
  syncCoordinator: LayoutSelectionGate;
}>();

const overlayRef = ref<HTMLDivElement | null>(null);
const renderEpoch = ref(0);
let rafId: number | null = null;
let unsubscribeRender: (() => void) | null = null;

function scheduleSync() {
  if (rafId !== null) cancelAnimationFrame(rafId);
  rafId = requestAnimationFrame(() => {
    rafId = null;
    const view = props.getView();
    const pagesContainer = props.getPagesContainer();
    const overlay = overlayRef.value;
    if (!view || !pagesContainer || !overlay) return;
    if (!props.syncCoordinator.isSafeToRender()) return;
    syncDecorations(view, pagesContainer, overlay, props.zoom);
  });
}

onMounted(() => {
  unsubscribeRender = props.syncCoordinator.onRender(() => {
    renderEpoch.value++;
  });
  scheduleSync();
});

onBeforeUnmount(() => {
  if (rafId !== null) {
    cancelAnimationFrame(rafId);
    rafId = null;
  }
  unsubscribeRender?.();
});

watch(
  () => [props.zoom, props.transactionVersion, renderEpoch.value],
  () => scheduleSync(),
);

function syncDecorations(
  view: EditorView,
  pagesContainer: HTMLElement,
  overlay: HTMLElement,
  zoom: number,
) {
  const decorations = collectDecorations(view.state);
  if (decorations.length === 0) {
    if (overlay.firstChild) overlay.replaceChildren();
    return;
  }

  const ctx = createRenderedDomContext(pagesContainer, zoom);
  const offset = ctx.getContainerOffset();
  const fragment = document.createDocumentFragment();

  for (const { decoration, from, to } of decorations) {
    if (from === to) {
      const dom = getWidgetDOM(decoration, view);
      if (!dom) continue;
      const coords = ctx.getCoordinatesForPosition(from);
      if (!coords) continue;
      const wrapper = document.createElement("div");
      wrapper.style.cssText =
        `position:absolute;left:${coords.x + offset.x}px;top:${coords.y + offset.y}px;` +
        `height:${coords.height}px;`;
      wrapper.appendChild(dom);
      fragment.appendChild(wrapper);
      continue;
    }

    const attrs = getDecorationAttrs(decoration);
    if (!attrs) continue;
    const rects = ctx.getRectsForRange(from, to);
    for (const rect of rects) {
      const el = document.createElement("div");
      for (const [name, value] of Object.entries(attrs)) {
        if (name === "nodeName") continue;
        el.setAttribute(name, value);
      }
      const baseStyle =
        `position:absolute;left:${rect.x + offset.x}px;top:${rect.y + offset.y}px;` +
        `width:${rect.width}px;height:${rect.height}px;`;
      el.style.cssText = baseStyle + (attrs["style"] ?? "");
      fragment.appendChild(el);
    }
  }

  overlay.replaceChildren(fragment);
}

type CollectedDecoration = {
  decoration: Decoration;
  from: number;
  to: number;
};

function collectDecorations(state: EditorState): CollectedDecoration[] {
  const out: CollectedDecoration[] = [];
  for (const plugin of state.plugins) {
    const decorationsFn = plugin.props.decorations;
    if (!decorationsFn) continue;
    const source = decorationsFn.call(plugin, state);
    if (!source) continue;
    // A `DecorationSource` walks down to its leaf `DecorationSet`s via
    // `forEachSet`. This overlay is a PORT-BLOCKED stub (see file header): read
    // the traversal reflectively and skip the source rather than throw if a
    // future/leaner prosemirror-view drops the method. The stub renders nothing
    // regardless, so failing soft here keeps it from crashing the editor.
    const forEachSet = readField(source, "forEachSet");
    if (typeof forEachSet !== "function") continue;
    forEachSet.call(source, (set: DecorationSet) => {
      set.find().forEach((decoration) => {
        const spec = readField(decoration, "spec");
        if (typeof spec === "object" && spec !== null && readField(spec, "noOverlay")) return;
        out.push({ decoration, from: decoration.from, to: decoration.to });
      });
    });
  }
  return out;
}

/**
 * Reflectively read an untyped field off a ProseMirror internal object (the
 * decoration's private `type`/`spec`) without an `as` cast. Returns `unknown`;
 * every call site narrows before use.
 */
function readField(target: object, key: string): unknown {
  return Reflect.get(target, key);
}

function getWidgetDOM(decoration: Decoration, view: EditorView): HTMLElement | null {
  const type = readField(decoration, "type");
  if (typeof type !== "object" || type === null) return null;
  const toDOM = readField(type, "toDOM");
  if (typeof toDOM === "function") {
    const dom: unknown = toDOM(view, () => decoration.from);
    return dom instanceof HTMLElement ? dom : null;
  }
  if (toDOM instanceof HTMLElement) {
    const clone = toDOM.cloneNode(true);
    return clone instanceof HTMLElement ? clone : null;
  }
  return null;
}

function getDecorationAttrs(decoration: Decoration): Record<string, string> | null {
  const type = readField(decoration, "type");
  if (typeof type !== "object" || type === null) return null;
  const attrs = readField(type, "attrs");
  if (typeof attrs !== "object" || attrs === null) return null;
  const out: Record<string, string> = {};
  for (const [name, value] of Object.entries(attrs)) {
    if (typeof value === "string") out[name] = value;
  }
  return out;
}

// PORT-BLOCKED stub — see the file header. A no-op RenderedDomContext so the
// component type-checks and mounts; it paints nothing until core's plugin-api
// (createRenderedDomContext) is ported.
type RenderedDomContextStub = {
  getContainerOffset(): { x: number; y: number };
  getCoordinatesForPosition(pos: number): { x: number; y: number; height: number } | null;
  getRectsForRange(
    from: number,
    to: number,
  ): Array<{ x: number; y: number; width: number; height: number }>;
};

function createRenderedDomContext(
  _pagesContainer: HTMLElement,
  _zoom: number,
): RenderedDomContextStub {
  return {
    getContainerOffset: () => ({ x: 0, y: 0 }),
    getCoordinatesForPosition: () => null,
    getRectsForRange: () => [],
  };
}
</script>
