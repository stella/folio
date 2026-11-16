/** Thin React binding for core's persistent header/footer editor manager. */
import {
  forwardRef,
  memo,
  useEffect,
  useImperativeHandle,
  useLayoutEffect,
  useRef,
} from "react";
import type { CSSProperties, NamedExoticComponent, RefAttributes } from "react";

import type { EditorView } from "prosemirror-view";

import {
  createHeaderFooterEditorManager,
  enumerateDocumentHeaderFooterParts,
  enumerateHeaderFooterParts,
} from "@stll/folio-core/controller/headerFooterEditorManager";
import type {
  HeaderFooterEditorManager,
  HeaderFooterPartKey,
  HeaderFooterPartKind,
} from "@stll/folio-core/controller/headerFooterEditorManager";
import type { HiddenEditorApi } from "@stll/folio-core/controller/hiddenEditorApi";
import type {
  CanonicalRefusalGround,
  EditorSession,
} from "@stll/folio-core/types/canonicalCapabilities";
import type {
  Document,
  HeaderFooter,
  StyleDefinitions,
  Theme,
} from "@stll/folio-core/types/document";
import type {
  DeepReadonly,
  ReadonlyDocument,
} from "@stll/folio-core/types/documentModel";

import "prosemirror-view/style/prosemirror.css";

export type HfPartKind = HeaderFooterPartKind;
export type HfPartKey = HeaderFooterPartKey;

export type HiddenHeaderFooterPMsRef = {
  getHostElement: () => HTMLElement | null;
  getView: (rId: string) => EditorView | null;
  listSlots: () => HfPartKey[];
};

export type HiddenHeaderFooterPMsProps = {
  /** The document model; the header and footer editors only read it. */
  document: ReadonlyDocument | null;
  readOnly?: boolean;
  experimentalSession?: EditorSession;
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
  styles?: DeepReadonly<StyleDefinitions> | null;
  theme?: DeepReadonly<Theme> | null;
  defaultTabStopTwips?: number | null;
  onTransaction?: (
    rId: string,
    kind: HfPartKind,
    view: EditorView,
    docChanged: boolean,
    selectionChanged: boolean,
  ) => void;
};

export const enumerateHfSlotsFromParts: ({
  headers,
  footers,
}: {
  headers: Map<string, HeaderFooter> | undefined;
  footers: Map<string, HeaderFooter> | undefined;
}) => HeaderFooterPartKey[] = enumerateHeaderFooterParts;
export const enumerateHfSlots: (
  document: Document | null,
) => HeaderFooterPartKey[] = enumerateDocumentHeaderFooterParts;

const HOST_STYLES: CSSProperties = {
  position: "fixed",
  left: -9999,
  top: 0,
  opacity: 0,
  zIndex: -1,
  pointerEvents: "none",
};

/* eslint-disable prefer-arrow-callback -- preserve the component name in React DevTools. */
export const HiddenHeaderFooterPMs: NamedExoticComponent<
  HiddenHeaderFooterPMsProps & RefAttributes<HiddenHeaderFooterPMsRef>
> = memo(
  forwardRef<HiddenHeaderFooterPMsRef, HiddenHeaderFooterPMsProps>(
    function HiddenHeaderFooterPMs(
      {
        document,
        readOnly,
        styles,
        theme,
        onTransaction,
        experimentalSession,
        getCanonicalApi,
        readCanonicalDocument,
        onSessionRefusal,
      },
      ref,
    ) {
      const hostRef = useRef<HTMLDivElement>(null);
      const documentRef = useRef(document);
      const readOnlyRef = useRef(readOnly);
      const stylesRef = useRef(styles);
      const themeRef = useRef(theme);
      const onTransactionRef = useRef(onTransaction);
      const canonicalApiRef = useRef(getCanonicalApi);
      const readCanonicalDocumentRef = useRef(readCanonicalDocument);
      const sessionRef = useRef(experimentalSession);
      const refusalRef = useRef(onSessionRefusal);

      const managerRef = useRef<HeaderFooterEditorManager | null>(null);
      useLayoutEffect(() => {
        documentRef.current = document;
        readOnlyRef.current = readOnly;
        stylesRef.current = styles;
        themeRef.current = theme;
        onTransactionRef.current = onTransaction;
        canonicalApiRef.current = getCanonicalApi;
        readCanonicalDocumentRef.current = readCanonicalDocument;
        sessionRef.current = experimentalSession;
        refusalRef.current = onSessionRefusal;
        managerRef.current ??= createHeaderFooterEditorManager({
          getHost: () => hostRef.current,
          getDocument: () => documentRef.current,
          getReadOnly: () => readOnlyRef.current ?? false,
          getStyles: () => stylesRef.current,
          getTheme: () => themeRef.current,
          getCanonicalApi: () => canonicalApiRef.current?.() ?? null,
          readCanonicalDocument: () => readCanonicalDocumentRef.current(),
          getExperimentalSession: () => sessionRef.current,
          onSessionRefusal: (reason, ground, error) =>
            refusalRef.current?.(reason, ground, error),
          onTransaction: ({
            rId,
            kind,
            view,
            docChanged,
            selectionChanged,
          }) => {
            onTransactionRef.current?.(
              rId,
              kind,
              view,
              docChanged,
              selectionChanged,
            );
          },
        });
        managerRef.current.sync();
      }, [
        document,
        readOnly,
        styles,
        theme,
        onTransaction,
        getCanonicalApi,
        readCanonicalDocument,
        experimentalSession,
        onSessionRefusal,
      ]);

      useEffect(
        () => () => {
          managerRef.current?.destroy();
        },
        [],
      );

      useImperativeHandle(
        ref,
        () => ({
          getHostElement: () => hostRef.current,
          getView: (rId) => managerRef.current?.getView(rId) ?? null,
          listSlots: () => managerRef.current?.listSlots() ?? [],
        }),
        [],
      );

      return (
        <div
          ref={hostRef}
          className="paged-editor__hidden-hf-pm"
          style={HOST_STYLES}
        />
      );
    },
  ),
);
/* eslint-enable prefer-arrow-callback */

HiddenHeaderFooterPMs.displayName = "HiddenHeaderFooterPMs";
