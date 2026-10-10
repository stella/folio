import type { CanonicalGap } from "@stll/folio-core/types/canonicalCapabilities";
/** Thin React binding for core's persistent header/footer editor manager. */

import { forwardRef, memo, useEffect, useLayoutEffect, useImperativeHandle, useRef } from "react";
import type { CSSProperties } from "react";

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
import type { Document, StyleDefinitions, Theme } from "@stll/folio-core/types/document";

import "prosemirror-view/style/prosemirror.css";

export type HfPartKind = HeaderFooterPartKind;
export type HfPartKey = HeaderFooterPartKey;

export type HiddenHeaderFooterPMsRef = {
  getHostElement: () => HTMLElement | null;
  getView: (rId: string) => EditorView | null;
  listSlots: () => HfPartKey[];
};

export type HiddenHeaderFooterPMsProps = {
  document: Document | null;
  experimentalSession?: "canonical";
  getCanonicalApi?: () => HiddenEditorApi | null;
  onSessionRefusal?: (reason: string, gap: CanonicalGap, error?: Error) => void;
  styles?: StyleDefinitions | null;
  theme?: Theme | null;
  defaultTabStopTwips?: number | null;
  onTransaction?: (
    rId: string,
    kind: HfPartKind,
    view: EditorView,
    docChanged: boolean,
    selectionChanged: boolean,
  ) => void;
};

export const enumerateHfSlotsFromParts = enumerateHeaderFooterParts;
export const enumerateHfSlots = enumerateDocumentHeaderFooterParts;

const HOST_STYLES: CSSProperties = {
  position: "fixed",
  left: -9999,
  top: 0,
  opacity: 0,
  zIndex: -1,
  pointerEvents: "none",
};

/* eslint-disable prefer-arrow-callback -- preserve the component name in React DevTools. */
export const HiddenHeaderFooterPMs = memo(
  forwardRef<HiddenHeaderFooterPMsRef, HiddenHeaderFooterPMsProps>(function HiddenHeaderFooterPMs(
    {
      document,
      styles,
      theme,
      onTransaction,
      experimentalSession,
      getCanonicalApi,
      onSessionRefusal,
    },
    ref,
  ) {
    const hostRef = useRef<HTMLDivElement>(null);
    const documentRef = useRef(document);
    const stylesRef = useRef(styles);
    const themeRef = useRef(theme);
    const onTransactionRef = useRef(onTransaction);
    const canonicalApiRef = useRef(getCanonicalApi);
    const sessionRef = useRef(experimentalSession);
    const refusalRef = useRef(onSessionRefusal);

    const managerRef = useRef<HeaderFooterEditorManager | null>(null);
    useLayoutEffect(() => {
      documentRef.current = document;
      stylesRef.current = styles;
      themeRef.current = theme;
      onTransactionRef.current = onTransaction;
      canonicalApiRef.current = getCanonicalApi;
      sessionRef.current = experimentalSession;
      refusalRef.current = onSessionRefusal;
      managerRef.current ??= createHeaderFooterEditorManager({
        getHost: () => hostRef.current,
        getDocument: () => documentRef.current,
        getStyles: () => stylesRef.current,
        getTheme: () => themeRef.current,
        getCanonicalApi: () => canonicalApiRef.current?.() ?? null,
        getExperimentalSession: () => sessionRef.current,
        onSessionRefusal: (reason, gap, error) => refusalRef.current?.(reason, gap, error),
        onTransaction: ({ rId, kind, view, docChanged, selectionChanged }) => {
          onTransactionRef.current?.(rId, kind, view, docChanged, selectionChanged);
        },
      });
    }, [
      document,
      styles,
      theme,
      onTransaction,
      getCanonicalApi,
      experimentalSession,
      onSessionRefusal,
    ]);

    useEffect(() => {
      managerRef.current?.sync();
    });

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

    return <div ref={hostRef} className="paged-editor__hidden-hf-pm" style={HOST_STYLES} />;
  }),
);
/* eslint-enable prefer-arrow-callback */

HiddenHeaderFooterPMs.displayName = "HiddenHeaderFooterPMs";
