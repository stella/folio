/** Owns browser font registration for one Vue editor instance. */

import { watch, type Ref } from "vue";

import type { FontDefinition } from "../components/DocxEditor/types";
import { loadEmbeddedFontFaces } from "../utils/embeddedFonts";
import { removeFontFaces } from "../utils/fontFaces";
import { loadHostFontFaces } from "../utils/hostFonts";

export type UseFontLifecycleOptions = {
  isReady: Ref<boolean>;
  getDocument: () => { originalBuffer?: ArrayBuffer | null } | null;
  fonts: () => ReadonlyArray<FontDefinition> | undefined;
  remeasure: () => void;
};

export type FontLifecycleDependencies = {
  loadEmbedded: (buffer: ArrayBuffer) => Promise<FontFace[]>;
  loadHost: (fonts: ReadonlyArray<FontDefinition> | undefined) => Promise<FontFace[]>;
  remove: (faces: readonly FontFace[]) => void;
};

const DEFAULT_DEPENDENCIES: FontLifecycleDependencies = {
  loadEmbedded: loadEmbeddedFontFaces,
  loadHost: loadHostFontFaces,
  remove: removeFontFaces,
};

/**
 * Register document-embedded and host-provided fonts for the lifetime of their
 * sources. Async results from a stale document or prop value are immediately
 * removed, preventing one editor load from leaking faces into the next.
 */
export function useFontLifecycle(
  options: UseFontLifecycleOptions,
  dependencies: FontLifecycleDependencies = DEFAULT_DEPENDENCIES,
): void {
  watch(
    () => (options.isReady.value ? (options.getDocument()?.originalBuffer ?? null) : null),
    (buffer, _previous, onCleanup) => {
      if (!buffer) {
        return;
      }

      let cancelled = false;
      let registered: FontFace[] = [];
      void dependencies.loadEmbedded(buffer).then((faces) => {
        if (cancelled) {
          dependencies.remove(faces);
          return undefined;
        }
        registered = faces;
        if (faces.length > 0) {
          options.remeasure();
        }
        return undefined;
      });

      onCleanup(() => {
        cancelled = true;
        dependencies.remove(registered);
      });
    },
    { immediate: true },
  );

  watch(
    options.fonts,
    (fonts, _previous, onCleanup) => {
      if (!fonts || fonts.length === 0) {
        return;
      }

      let cancelled = false;
      let registered: FontFace[] = [];
      void dependencies.loadHost(fonts).then((faces) => {
        if (cancelled) {
          dependencies.remove(faces);
          return undefined;
        }
        registered = faces;
        if (faces.length > 0) {
          options.remeasure();
        }
        return undefined;
      });

      onCleanup(() => {
        cancelled = true;
        if (registered.length === 0) {
          return;
        }
        dependencies.remove(registered);
        options.remeasure();
      });
    },
    { immediate: true },
  );
}
