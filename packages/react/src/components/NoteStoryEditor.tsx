/** Thin React surface over core's persistent note-story editor manager. */

import { forwardRef, useEffect, useImperativeHandle, useRef, useState } from "react";
import type { CSSProperties } from "react";
import { useTranslations } from "use-intl";

import type { EditorView } from "prosemirror-view";
import type { Plugin } from "prosemirror-state";

import {
  createNoteEditorManager,
  type NoteEditorManager,
  type NoteStoryKey,
} from "@stll/folio-core/controller/noteEditorManager";
import { setSuggestionMode } from "@stll/folio-core/prosemirror/plugins/suggestionMode";
import type { Document, StyleDefinitions, Theme } from "@stll/folio-core/types/document";

import "prosemirror-view/style/prosemirror.css";

export type NoteStoryEditorRef = {
  close: () => void;
  getActiveView: () => EditorView | null;
  open: (story: NoteStoryKey) => void;
  snapshotDocument: (document: Document) => Document;
};

export type NoteStoryEditorProps = {
  document: Document | null;
  onActiveChange: (story: NoteStoryKey | null) => void;
  onDocumentChange: (document: Document) => void;
  onStoryChange: (view: EditorView, docChanged: boolean, selectionChanged: boolean) => void;
  plugins?: Plugin[];
  suggestionAuthor?: string;
  suggestionModeActive?: boolean;
  styles?: StyleDefinitions | null;
  theme?: Theme | null;
};

const panelStyle: CSSProperties = {
  position: "absolute",
  right: 24,
  bottom: 24,
  width: "min(560px, calc(100% - 48px))",
  maxHeight: "45%",
  overflow: "auto",
  zIndex: 30,
  background: "var(--doc-canvas, #fff)",
  color: "var(--doc-canvas-text, #000)",
  border: "1px solid var(--doc-border, rgba(0, 0, 0, 0.18))",
  borderRadius: 6,
  boxShadow: "0 8px 28px rgba(0, 0, 0, 0.18)",
};

const headerStyle: CSSProperties = {
  position: "sticky",
  top: 0,
  zIndex: 1,
  display: "flex",
  alignItems: "center",
  justifyContent: "space-between",
  padding: "8px 12px",
  background: "var(--doc-canvas, #fff)",
  borderBottom: "1px solid var(--doc-border, rgba(0, 0, 0, 0.12))",
  fontSize: 12,
  fontWeight: 600,
};

const hostStyle: CSSProperties = { padding: "12px 16px", minHeight: 72 };
const EMPTY_PLUGINS: Plugin[] = [];

/* eslint-disable prefer-arrow-callback -- preserve the component name in React DevTools. */
export const NoteStoryEditor = forwardRef<NoteStoryEditorRef, NoteStoryEditorProps>(
  function NoteStoryEditor(
    {
      document,
      onActiveChange,
      onDocumentChange,
      onStoryChange,
      plugins,
      styles,
      suggestionAuthor,
      suggestionModeActive,
      theme,
    },
    ref,
  ) {
    const t = useTranslations("folio");
    const resolvedPlugins = plugins ?? EMPTY_PLUGINS;
    const isSuggestionModeActive = suggestionModeActive ?? false;
    const hostRef = useRef<HTMLDivElement>(null);
    const documentRef = useRef(document);
    const stylesRef = useRef(styles);
    const themeRef = useRef(theme);
    const pluginsRef = useRef(resolvedPlugins);
    const onDocumentChangeRef = useRef(onDocumentChange);
    const onActiveChangeRef = useRef(onActiveChange);
    const onStoryChangeRef = useRef(onStoryChange);
    const [active, setActive] = useState<NoteStoryKey | null>(null);
    const managerRef = useRef<NoteEditorManager | null>(null);

    const footnotes = document?.package.footnotes;
    const endnotes = document?.package.endnotes;
    useEffect(() => {
      documentRef.current = document;
      stylesRef.current = styles;
      themeRef.current = theme;
      pluginsRef.current = resolvedPlugins;
      onDocumentChangeRef.current = onDocumentChange;
      onActiveChangeRef.current = onActiveChange;
      onStoryChangeRef.current = onStoryChange;
      if (managerRef.current === null) {
        managerRef.current = createNoteEditorManager({
          getDocument: () => documentRef.current,
          getHost: () => hostRef.current,
          getPlugins: () => pluginsRef.current,
          getStyles: () => stylesRef.current,
          getTheme: () => themeRef.current,
          onTransaction: ({ docChanged, selectionChanged, view }) => {
            if (docChanged) {
              const current = documentRef.current;
              if (current) {
                onDocumentChangeRef.current(
                  managerRef.current?.snapshotDocument(current) ?? current,
                );
              }
            }
            onStoryChangeRef.current(view, docChanged, selectionChanged);
          },
        });
      }
      const manager = managerRef.current;
      manager.sync();
      if (active && !manager?.getActive()) {
        setActive(null);
        onActiveChangeRef.current(null);
      }
    }, [
      active,
      document,
      endnotes,
      footnotes,
      onActiveChange,
      onDocumentChange,
      onStoryChange,
      resolvedPlugins,
      styles,
      theme,
    ]);

    useEffect(() => {
      const manager = managerRef.current;
      if (!manager) return;
      for (const story of manager.listStories()) {
        const view = manager.getView(story);
        if (view) {
          setSuggestionMode(isSuggestionModeActive, view.state, view.dispatch, suggestionAuthor);
        }
      }
    }, [isSuggestionModeActive, suggestionAuthor]);

    useEffect(
      () => () => {
        managerRef.current?.destroy();
      },
      [],
    );

    const closeEditor = (): void => {
      managerRef.current?.activate(null);
      setActive(null);
      onActiveChangeRef.current(null);
    };

    useImperativeHandle(ref, () => ({
      close: closeEditor,
      getActiveView: () => {
        const story = managerRef.current?.getActive();
        return story ? (managerRef.current?.getView(story) ?? null) : null;
      },
      open: (story) => {
        const view = managerRef.current?.activate(story);
        if (!view) return;
        setSuggestionMode(isSuggestionModeActive, view.state, view.dispatch, suggestionAuthor);
        setActive(story);
        onActiveChangeRef.current(story);
        requestAnimationFrame(() => view.focus());
      },
      snapshotDocument: (current) => managerRef.current?.snapshotDocument(current) ?? current,
    }));

    const title = active
      ? t(
          active.kind === "footnote"
            ? "dialogs.footnoteProperties.footnotes"
            : "dialogs.footnoteProperties.endnotes",
        )
      : "";

    return (
      <aside
        style={{ ...panelStyle, display: active ? "block" : "none" }}
        aria-label={title}
        onKeyDown={(event) => {
          if (event.key !== "Escape") return;
          event.preventDefault();
          event.stopPropagation();
          closeEditor();
        }}
      >
        <div style={headerStyle}>
          <span>{title}</span>
          <button
            type="button"
            aria-label={t("common.closeDialog")}
            onClick={closeEditor}
            style={{ border: 0, background: "transparent", color: "inherit", cursor: "pointer" }}
          >
            ×
          </button>
        </div>
        <div ref={hostRef} style={hostStyle} />
      </aside>
    );
  },
);
/* eslint-enable prefer-arrow-callback */

NoteStoryEditor.displayName = "NoteStoryEditor";
