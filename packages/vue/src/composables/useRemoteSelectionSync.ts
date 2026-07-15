import { onScopeDispose, watch, type Ref } from "vue";

import type { HiddenProseMirrorRemoteSelection } from "@stll/folio-core/controller/hiddenEditorManager";
import type { LayoutSelectionGate } from "@stll/folio-core/paged-layout/LayoutSelectionGate";
import { RemoteSelectionOverlay } from "@stll/folio-core/render-dom/RemoteSelectionOverlay";

import { Z_INDEX } from "../styles/zIndex";

export type UseRemoteSelectionSyncOptions = {
  pagesRef: Ref<HTMLElement | null>;
  remoteSelections: Ref<HiddenProseMirrorRemoteSelection[]>;
  syncCoordinator: LayoutSelectionGate;
  zoom: Ref<number>;
};

export const useRemoteSelectionSync = (options: UseRemoteSelectionSyncOptions): void => {
  const overlay = new RemoteSelectionOverlay();
  let animationFrame: number | null = null;

  const clear = (): void => {
    if (animationFrame !== null) {
      cancelAnimationFrame(animationFrame);
      animationFrame = null;
    }
    const pagesContainer = options.pagesRef.value;
    if (pagesContainer) {
      overlay.clear(pagesContainer);
    }
  };

  const sync = (): void => {
    if (animationFrame !== null) {
      cancelAnimationFrame(animationFrame);
    }
    animationFrame = requestAnimationFrame(() => {
      animationFrame = null;
      if (!options.syncCoordinator.isSafeToRender()) {
        return;
      }
      const pagesContainer = options.pagesRef.value;
      if (!pagesContainer) {
        return;
      }
      overlay.sync({
        pagesContainer,
        selections: options.remoteSelections.value,
        zoom: options.zoom.value,
        zIndex: Z_INDEX.remoteSelection,
      });
    });
  };

  const unsubscribeRender = options.syncCoordinator.onRender(sync);
  watch([options.pagesRef, options.remoteSelections, options.zoom], sync, { flush: "post" });
  onScopeDispose(() => {
    unsubscribeRender();
    clear();
  });
};
