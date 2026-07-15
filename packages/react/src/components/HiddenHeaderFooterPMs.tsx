/** Thin React binding for core's persistent header/footer editor manager. */

import { forwardRef, memo, useEffect, useImperativeHandle, useRef } from "react";
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
import type { Document, StyleDefinitions, Theme } from "@stll/folio-core/types/document";

import "prosemirror-view/style/prosemirror.css";

export type HfPartKind = HeaderFooterPartKind;
export type HfPartKey = HeaderFooterPartKey;

export type HiddenHeaderFooterPMsRef = {
  getView: (rId: string) => EditorView | null;
  listSlots: () => HfPartKey[];
};

export type HiddenHeaderFooterPMsProps = {
  document: Document | null;
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
    { document, styles, theme, onTransaction },
    ref,
  ) {
    const hostRef = useRef<HTMLDivElement>(null);
    const documentRef = useRef(document);
    const stylesRef = useRef(styles);
    const themeRef = useRef(theme);
    const onTransactionRef = useRef(onTransaction);
    documentRef.current = document;
    stylesRef.current = styles;
    themeRef.current = theme;
    onTransactionRef.current = onTransaction;

    const managerRef = useRef<HeaderFooterEditorManager | null>(null);
    managerRef.current ??= createHeaderFooterEditorManager({
      getHost: () => hostRef.current,
      getDocument: () => documentRef.current,
      getStyles: () => stylesRef.current,
      getTheme: () => themeRef.current,
      onTransaction: ({ rId, kind, view, docChanged, selectionChanged }) => {
        onTransactionRef.current?.(rId, kind, view, docChanged, selectionChanged);
      },
    });

    const headers = document?.package.headers;
    const footers = document?.package.footers;
    useEffect(() => {
      managerRef.current?.sync();
    }, [headers, footers, styles, theme]);

    useEffect(
      () => () => {
        managerRef.current?.destroy();
      },
      [],
    );

    useImperativeHandle(
      ref,
      () => ({
        getView: (rId) => managerRef.current?.getView(rId) ?? null,
        listSlots: () => managerRef.current?.listSlots() ?? [],
      }),
      [],
    );

    return <div ref={hostRef} style={HOST_STYLES} />;
  }),
);
/* eslint-enable prefer-arrow-callback */

HiddenHeaderFooterPMs.displayName = "HiddenHeaderFooterPMs";
