/**
 * Vue `renderAsync` — imperatively mount a DocxEditor into a DOM element.
 *
 * Mirrors the React adapter's `renderAsync` (`packages/react/src/renderAsync.tsx`):
 * returns a framework-agnostic {@link EditorHandle} extended with zoom and scroll
 * helpers. Resolves once the document has parsed (the editor's first `ready`
 * event), rejects if `error` fires before that milestone.
 *
 * ```ts
 * import { renderAsync } from "@stll/folio-vue";
 * import "@stll/folio-vue/standalone.css";
 *
 * const editor = await renderAsync(buffer, document.getElementById("host")!, {
 *   readOnly: false,
 *   showToolbar: true,
 * });
 *
 * const blob = await editor.save();
 * editor.destroy();
 * ```
 */

import { createApp, h, ref, type App, type MaybeRefOrGetter } from "vue";

import type { Document } from "@stll/folio-core/types/document";
import type { DocxInput } from "@stll/folio-core/utils/docxInput";

import DocxEditor from "./components/DocxEditor.vue";
import type { DocxEditorProps, DocxEditorRef } from "./components/DocxEditor/types";
import { colorModePlugin, defaultColorMode, type ColorMode } from "./composables/useColorMode";
import { i18nPlugin } from "./i18n";

/** Framework-agnostic handle for an imperatively mounted editor instance. */
export type EditorHandle = {
  /** Save the document and return the DOCX as a Blob. */
  save: () => Promise<Blob | null>;
  /** Get the current parsed document model. */
  getDocument: () => Document | null;
  /** Focus the editor. */
  focus: () => void;
  /** Unmount the editor and clean up. */
  destroy: () => void;
};

/** Vue-specific handle with zoom and scroll helpers. */
export type DocxEditorHandle = EditorHandle & {
  /** Set zoom level (1.0 = 100%). */
  setZoom: (zoom: number) => void;
  /** Scroll the visible pages to a raw ProseMirror document position. */
  scrollToPosition: (pmPos: number) => void;
  /** Scroll the visible pages to a 1-indexed page number. */
  scrollToPage: (pageNumber: number) => void;
};

/** Options for {@link renderAsync}. */
export type RenderAsyncOptions = Omit<DocxEditorProps, "documentBuffer" | "document"> & {
  /** BCP-47 locale for bundled folio UI strings (default: `en`). */
  locale?: string;
  /** Editor chrome color mode (default: `light`). */
  colorMode?: MaybeRefOrGetter<ColorMode>;
};

/**
 * Render a DOCX editor into a container element using Vue.
 *
 * Resolves once the document has parsed and the editor emits `ready`.
 * Rejects if `error` runs before that milestone.
 */
export const renderAsync = (
  input: DocxInput,
  container: HTMLElement,
  options: RenderAsyncOptions = {},
): Promise<DocxEditorHandle> => {
  const { colorMode = defaultColorMode, locale = "en", ...editorOptions } = options;

  return new Promise<DocxEditorHandle>((resolve, reject) => {
    const editorRef = ref<DocxEditorRef | null>(null);
    let app: App | null = null;
    let settled = false;

    const handle: DocxEditorHandle = {
      save: async () => {
        const buffer = await (editorRef.value?.save() ?? Promise.resolve(null));
        if (!buffer) {
          return null;
        }
        return new Blob([buffer], {
          type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        });
      },
      getDocument: () => editorRef.value?.getDocument() ?? null,
      focus: () => {
        editorRef.value?.focus();
      },
      setZoom: (zoom) => {
        editorRef.value?.setZoom(zoom);
      },
      scrollToPosition: (pmPos) => {
        editorRef.value?.getEditorRef()?.scrollToPosition(pmPos);
      },
      scrollToPage: (pageNumber) => {
        editorRef.value?.scrollToPage(pageNumber);
      },
      destroy: () => {
        if (!settled) {
          settled = true;
          reject(new Error("Editor was destroyed before mounting completed."));
        }
        app?.unmount();
        app = null;
      },
    };

    app = createApp({
      setup() {
        return () =>
          h(DocxEditor, {
            ...editorOptions,
            documentBuffer: input,
            ref: editorRef,
            onReady: () => {
              if (!settled) {
                settled = true;
                resolve(handle);
              }
            },
            onError: (error: Error) => {
              editorOptions.onError?.(error);
              if (!settled) {
                settled = true;
                app?.unmount();
                app = null;
                reject(error);
              }
            },
          });
      },
    });

    app.use(i18nPlugin, locale);
    app.use(colorModePlugin, colorMode);
    app.mount(container);
  });
};
