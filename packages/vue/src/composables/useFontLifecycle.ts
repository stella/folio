/** Owns browser font registration for one Vue editor instance. */

import { watch, type Ref } from "vue";
import { setEmbeddedFontFamilyMap } from "@stll/folio-core/utils/fontResolver";

import type { FontDefinition } from "../components/DocxEditor/types";
import { loadEmbeddedFontFaces, type LoadedEmbeddedFonts } from "../utils/embeddedFonts";
import { removeFontFaces } from "../utils/fontFaces";
import { loadHostFontFaces } from "../utils/hostFonts";

export type UseFontLifecycleOptions = {
  isReady: Ref<boolean>;
  getDocument: () => { originalBuffer?: ArrayBuffer | null } | null;
  fonts: () => ReadonlyArray<FontDefinition> | undefined;
  remeasure: () => void;
};

export type FontLifecycleDependencies = {
  loadEmbedded: (buffer: ArrayBuffer) => Promise<LoadedEmbeddedFonts>;
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
      void dependencies.loadEmbedded(buffer).then(({ faces, familyMap }) => {
        if (cancelled) {
          dependencies.remove(faces);
          return undefined;
        }
        registered = faces;
        if (faces.length === 0 && familyMap.size === 0) {
          return undefined;
        }
        // Activate this document's original→scoped embedded-font family map
        // before `options.remeasure()` invalidates the resolved-font cache,
        // so every run that resolves after this point picks up the scoped
        // family instead of the raw DOCX name — which is never registered on
        // `document.fonts` (see `@stll/folio-core/fonts/embeddedFonts`).
        setEmbeddedFontFamilyMap(familyMap.size > 0 ? familyMap : null);
        options.remeasure();
        return undefined;
      });

      onCleanup(() => {
        cancelled = true;
        dependencies.remove(registered);
        setEmbeddedFontFamilyMap(null);
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
