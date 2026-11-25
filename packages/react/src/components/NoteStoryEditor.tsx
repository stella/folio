import {
  forwardRef,
  useEffect,
  useImperativeHandle,
  useRef,
  useState,
} from "react";
import type {
  CSSProperties,
  ForwardRefExoticComponent,
  RefAttributes,
} from "react";

/** Thin React surface over core's persistent note-story editor manager. */
import type { EditorView } from "prosemirror-view";
import { useTranslations } from "use-intl";

import type { HiddenEditorApi } from "@stll/folio-core/controller/hiddenEditorApi";
import {
  createNoteEditorManager,
  type NoteEditorManager,
  type NoteStoryKey,
} from "@stll/folio-core/controller/noteEditorManager";
import type { CanonicalRefusalGround } from "@stll/folio-core/types/canonicalCapabilities";
import type { StyleDefinitions } from "@stll/folio-core/types/document";
import type {
  DeepReadonly,
  ReadonlyDocument,
} from "@stll/folio-core/types/documentModel";

import "prosemirror-view/style/prosemirror.css";

export type NoteStoryEditorRef = {
  close: () => void;
  getActiveView: () => EditorView | null;
  open: (story: NoteStoryKey) => void;
  /** Brings every mounted note editor to the session's document. */
  sync: () => void;
};

export type NoteStoryEditorProps = {
  /** The document model; the note editors only read it. */
  document: ReadonlyDocument | null;
  getCanonicalApi?: () => HiddenEditorApi | null;
  /**
   * The canonical session's own document, read without a copy; null outside
   * a session. The editor's story managers only read it.
   */
  readCanonicalDocument: () => ReadonlyDocument | null;
  onSessionRefusal?: (
    reason: string,
    ground: CanonicalRefusalGround,
    error?: Error,
  ) => void;
  onActiveChange: (story: NoteStoryKey | null) => void;
  /** The model a note edit left, as it is typed: read it, never write it. */
  onDocumentChange: (model: ReadonlyDocument) => void;
  onStoryChange: (
    view: EditorView,
    docChanged: boolean,
    selectionChanged: boolean,
  ) => void;
  styles?: DeepReadonly<StyleDefinitions> | null;
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

/* eslint-disable prefer-arrow-callback -- preserve the component name in React DevTools. */
const NoteStoryEditorComponent = forwardRef<
  NoteStoryEditorRef,
  NoteStoryEditorProps
>(function NoteStoryEditor(
  {
    document,
    getCanonicalApi,
    readCanonicalDocument,
    onSessionRefusal,
    onActiveChange,
    onDocumentChange,
    onStoryChange,
    styles,
  },
  ref,
) {
  const t = useTranslations("folio");
  const hostRef = useRef<HTMLDivElement>(null);
  const canonicalApiRef = useRef(getCanonicalApi);
  const readCanonicalDocumentRef = useRef(readCanonicalDocument);
  const refusalRef = useRef(onSessionRefusal);
  const stylesRef = useRef(styles);
  const onDocumentChangeRef = useRef(onDocumentChange);
  const onActiveChangeRef = useRef(onActiveChange);
  const onStoryChangeRef = useRef(onStoryChange);
  const [active, setActive] = useState<NoteStoryKey | null>(null);
  const managerRef = useRef<NoteEditorManager | null>(null);

  useEffect(() => {
    canonicalApiRef.current = getCanonicalApi;
    readCanonicalDocumentRef.current = readCanonicalDocument;
    refusalRef.current = onSessionRefusal;
    stylesRef.current = styles;
    onDocumentChangeRef.current = onDocumentChange;
    onActiveChangeRef.current = onActiveChange;
    onStoryChangeRef.current = onStoryChange;
    if (managerRef.current === null) {
      managerRef.current = createNoteEditorManager({
        getHost: () => hostRef.current,
        getStyles: () => stylesRef.current,
        getCanonicalApi: () => canonicalApiRef.current?.() ?? null,
        readCanonicalDocument: () => readCanonicalDocumentRef.current(),
        onSessionRefusal: (reason, ground, error) =>
          refusalRef.current?.(reason, ground, error),
        onTransaction: ({ docChanged, selectionChanged, view }) => {
          if (docChanged) {
            // The session's own document, never a copy; the host's copy is
            // made once the edit settles.
            const model = readCanonicalDocumentRef.current();
            if (model) onDocumentChangeRef.current(model);
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
    getCanonicalApi,
    readCanonicalDocument,
    onActiveChange,
    onDocumentChange,
    onSessionRefusal,
    onStoryChange,
    styles,
  ]);

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
      setActive(story);
      onActiveChangeRef.current(story);
      requestAnimationFrame(() => view.focus());
    },
    sync: () => managerRef.current?.sync(),
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
          style={{
            border: 0,
            background: "transparent",
            color: "inherit",
            cursor: "pointer",
          }}
        >
          ×
        </button>
      </div>
      <div ref={hostRef} style={hostStyle} />
    </aside>
  );
});
/* eslint-enable prefer-arrow-callback */

export const NoteStoryEditor: ForwardRefExoticComponent<
  NoteStoryEditorProps & RefAttributes<NoteStoryEditorRef>
> = NoteStoryEditorComponent;

NoteStoryEditor.displayName = "NoteStoryEditor";
