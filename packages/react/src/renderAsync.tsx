/**
 * Imperative API for mounting a DOCX editor into a DOM element.
 *
 * Returns a framework-agnostic {@link EditorHandle} extended with zoom and
 * scroll helpers. Wraps the editor in `IntlProvider` so `use-intl` works
 * without extra host setup.
 *
 * ```ts
 * import { renderAsync } from "@stll/folio-react";
 * import "@stll/folio-react/standalone.css";
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

import React from "react";
import { createRoot, type Root } from "react-dom/client";
import { IntlProvider } from "use-intl";

import type { Document } from "@stll/folio-core/types/document";
import type { DocxInput } from "@stll/folio-core/utils/docxInput";

import { DocxEditor } from "./components/DocxEditor";
import type { DocxEditorProps, DocxEditorRef } from "./components/DocxEditor.props";
import { getFolioMessages } from "./i18n/messages";

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

/** React-specific handle with zoom and scroll helpers. */
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
};

/**
 * Render a DOCX editor into a container element.
 *
 * Resolves once the document has parsed and the first `onChange` fires.
 * Rejects if `onError` runs before that milestone.
 */
export const renderAsync = (
  input: DocxInput,
  container: HTMLElement,
  options: RenderAsyncOptions = {},
): Promise<DocxEditorHandle> => {
  const { locale = "en", ...editorOptions } = options;

  return new Promise<DocxEditorHandle>((resolve, reject) => {
    const ref = React.createRef<DocxEditorRef>();
    let root: Root | null = null;

    try {
      root = createRoot(container);
    } catch (error) {
      reject(error);
      return;
    }

    const handle: DocxEditorHandle = {
      save: async () => {
        const buffer = await (ref.current?.save() ?? Promise.resolve(null));
        if (!buffer) {
          return null;
        }
        return new Blob([buffer], {
          type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        });
      },
      getDocument: () => ref.current?.getDocument() ?? null,
      focus: () => {
        ref.current?.focus();
      },
      setZoom: (zoom) => {
        ref.current?.setZoom(zoom);
      },
      scrollToPosition: (pmPos) => {
        ref.current?.getEditorRef()?.scrollToPosition(pmPos);
      },
      scrollToPage: (pageNumber) => {
        ref.current?.scrollToPage(pageNumber);
      },
      destroy: () => {
        root?.unmount();
        root = null;
      },
    };

    let settled = false;

    const element = (
      <IntlProvider locale={locale} messages={getFolioMessages(locale)}>
        <DocxEditor
          {...editorOptions}
          documentBuffer={input}
          onError={(error) => {
            editorOptions["onError"]?.(error);
            if (!settled) {
              settled = true;
              reject(error);
            }
          }}
          onChange={(doc) => {
            editorOptions["onChange"]?.(doc);
            if (!settled) {
              settled = true;
              resolve(handle);
            }
          }}
          ref={ref}
        />
      </IntlProvider>
    );

    root.render(element);
  });
};
