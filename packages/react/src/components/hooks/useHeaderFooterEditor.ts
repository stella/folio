/**
 * Hook encapsulating header/footer editing state, content resolution, and
 * mutation callbacks extracted from DocxEditor.
 *
 * Thin React binding around the framework-agnostic header/footer operations in
 * `@stll/folio-core/utils/headerFooter`: the core module resolves the rendered
 * slots and builds the create/save/remove Documents; this hook keeps the React
 * glue — the edit-position state, the memoised resolution, the host
 * `pushDocument` routing, and reading the live hidden HF PM doc at save time.
 */

import { useCallback, useMemo, useRef, useState } from "react";

import type { EditorView } from "prosemirror-view";

import { proseDocToBlocks } from "@stll/folio-core/prosemirror/conversion/fromProseDoc";
import type {
  Document,
  HeaderFooter,
  SectionProperties,
} from "@stll/folio-core/types/document";
import {
  createEmptyHeaderFooter,
  pickActiveHeaderFooterRId,
  removeHeaderFooter,
  resolveEffectiveSectionProperties,
  resolveHeaderFooterContent,
  saveHeaderFooterContent,
} from "@stll/folio-core/utils/headerFooter";
import type { UseHistoryReturn } from "../../hooks/useHistory";

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

type UseHeaderFooterEditorParams = {
  history: UseHistoryReturn<Document | null>;
  pushDocument: (document: Document) => Document;
  /**
   * Look up the persistent hidden HF EditorView for an rId. Returns null
   * when the view isn't mounted (e.g. the chrome unmounted before the
   * close path ran). Called at close time so the save path can read live
   * PM state and serialise it into the new HeaderFooter without mutating
   * the current history.state.
   */
  getHfView: (rId: string) => EditorView | null;
};

type UseHeaderFooterEditorReturn = {
  /** Which header/footer area is currently being edited, or null */
  hfEditPosition: "header" | "footer" | null;
  setHfEditPosition: (pos: "header" | "footer" | null) => void;
  /** Whether the current HF edit targets the first page */
  hfEditIsFirstPage: boolean;

  /** Resolved header/footer content for the active document */
  headerContent: HeaderFooter | null;
  footerContent: HeaderFooter | null;
  firstPageHeaderContent: HeaderFooter | null;
  firstPageFooterContent: HeaderFooter | null;
  hasTitlePg: boolean;

  /**
   * Relationship ids for the *displayed* H/F slots — same values used to
   * route save/remove (Codex PR #258). The persistent hidden HF PM model
   * looks views up by these rIds; the painter emits them as `data-rid` so
   * the pointer pipeline can route clicks back to the matching view.
   */
  activeHeaderRId: string | null;
  activeFooterRId: string | null;
  activeFirstHeaderRId: string | null;
  activeFirstFooterRId: string | null;

  /** Section properties with titlePg merged from inline sections */
  effectiveSectionProperties: SectionProperties | undefined;

  /** Open the inline HF editor on double-click; creates an empty HF if needed */
  handleHeaderFooterDoubleClick: (position: "header" | "footer", pageNumber?: number) => void;
  /**
   * Snapshot the current HF state into a new Document via pushDocument so
   * the edit session lands in undo history. Reads `content` from the
   * active HF PM (via getHfView) at call time; nothing in `history.state`
   * is mutated until the new Document is pushed.
   */
  handleHeaderFooterSave: () => void;
  /** Save and close HF editor when the body is clicked */
  handleBodyClick: () => void;
  /** Remove the active header/footer from the document */
  handleRemoveHeaderFooter: () => void;
};

// ---------------------------------------------------------------------------
// Hook
// ---------------------------------------------------------------------------

export const useHeaderFooterEditor = ({
  history,
  pushDocument,
  getHfView,
}: UseHeaderFooterEditorParams): UseHeaderFooterEditorReturn => {
  // -------------------------------------------------------------------------
  // State
  // -------------------------------------------------------------------------

  const [hfEditPosition, setHfEditPosition] = useState<"header" | "footer" | null>(null);
  const [hfEditIsFirstPage, setHfEditIsFirstPage] = useState(false);

  // -------------------------------------------------------------------------
  // Resolved header/footer content
  // -------------------------------------------------------------------------

  const rawResolution = useMemo(
    () => resolveHeaderFooterContent(history.state?.package),
    [history.state],
  );
  // `history.state` is transiently null while the document buffer is parsed and
  // during relayout. Falling through to the empty resolution on those renders
  // blanks the active H/F rIds, which breaks edit focus (getActiveView returns
  // null, so typing lands in the body) and flickers the rendered chrome. Hold
  // the last resolved structure and reuse it until a real document state
  // arrives again.
  const lastResolution = useRef(rawResolution);
  if (history.state?.package) {
    lastResolution.current = rawResolution;
  }
  const resolution = history.state?.package ? rawResolution : lastResolution.current;
  const {
    headerContent,
    footerContent,
    firstPageHeaderContent,
    firstPageFooterContent,
    hasTitlePg,
    activeHeaderRId,
    activeFooterRId,
    activeFirstHeaderRId,
    activeFirstFooterRId,
  } = resolution;

  // -------------------------------------------------------------------------
  // Effective section properties (titlePg merged)
  // -------------------------------------------------------------------------

  const effectiveSectionProperties = useMemo(
    () => resolveEffectiveSectionProperties(history.state?.package.document, hasTitlePg),
    [history.state?.package.document, hasTitlePg],
  );

  // -------------------------------------------------------------------------
  // Callbacks
  // -------------------------------------------------------------------------

  const handleHeaderFooterDoubleClick = useCallback(
    (position: "header" | "footer", pageNumber?: number) => {
      const isFirstPage = hasTitlePg && (pageNumber ?? 1) === 1;
      let hf = position === "header" ? headerContent : footerContent;
      if (isFirstPage) {
        hf = position === "header" ? firstPageHeaderContent : firstPageFooterContent;
      }
      setHfEditIsFirstPage(isFirstPage);
      if (hf) {
        setHfEditPosition(position);
        return;
      }

      // Create an empty header/footer for docs that don't have one yet.
      if (!history.state) {
        return;
      }
      const newDoc = createEmptyHeaderFooter(history.state, position, isFirstPage);
      if (!newDoc) {
        return;
      }
      pushDocument(newDoc);
      setHfEditPosition(position);
    },
    [
      headerContent,
      footerContent,
      firstPageHeaderContent,
      firstPageFooterContent,
      hasTitlePg,
      history,
      pushDocument,
    ],
  );

  const handleHeaderFooterSave = useCallback(() => {
    if (!hfEditPosition || !history.state?.package) {
      setHfEditPosition(null);
      return;
    }

    // Resolve the rId via the SAME resolution that picks the displayed H/F —
    // for multi-section docs (NVCA-style title-page + signature sections) this
    // is the rendered rId, not finalSectionProperties' rId. Codex PR #258.
    const activeRId = pickActiveHeaderFooterRId(resolution, hfEditPosition, hfEditIsFirstPage);
    const view = activeRId ? getHfView(activeRId) : null;
    if (activeRId && view) {
      // Read fresh blocks from PM state — HiddenHeaderFooterPMs no longer
      // mutates `existing.content` in place (Codex #487 P1 re-fixed), so the
      // pre-edit snapshot stays intact and undo can step back to it.
      const newDoc = saveHeaderFooterContent({
        document: history.state,
        position: hfEditPosition,
        isFirstPage: hfEditIsFirstPage,
        activeRId,
        blocks: proseDocToBlocks(view.state.doc),
      });
      if (newDoc) {
        pushDocument(newDoc);
      }
    }

    setHfEditPosition(null);
  }, [hfEditPosition, hfEditIsFirstPage, resolution, history, pushDocument, getHfView]);

  const handleBodyClick = useCallback(() => {
    if (!hfEditPosition) {
      return;
    }
    // HF content is kept current by the HF PM's in-place sync
    // (HiddenHeaderFooterPMs.dispatchTransaction); the close path just needs to
    // publish the current state as a history snapshot.
    handleHeaderFooterSave();
  }, [hfEditPosition, handleHeaderFooterSave]);

  const handleRemoveHeaderFooter = useCallback(() => {
    if (!hfEditPosition || !history.state?.package) {
      setHfEditPosition(null);
      return;
    }

    // Same active-rId resolution as save: target the rId actually rendered, not
    // whatever lives in `finalSectionProperties` (Codex PR #258).
    const activeRId = pickActiveHeaderFooterRId(resolution, hfEditPosition, hfEditIsFirstPage);
    if (activeRId) {
      pushDocument(
        removeHeaderFooter({ document: history.state, position: hfEditPosition, activeRId }),
      );
    }

    setHfEditPosition(null);
  }, [hfEditPosition, hfEditIsFirstPage, resolution, history, pushDocument]);

  return {
    hfEditPosition,
    setHfEditPosition,
    hfEditIsFirstPage,
    headerContent,
    footerContent,
    firstPageHeaderContent,
    firstPageFooterContent,
    hasTitlePg,
    activeHeaderRId,
    activeFooterRId,
    activeFirstHeaderRId,
    activeFirstFooterRId,
    effectiveSectionProperties,
    handleHeaderFooterDoubleClick,
    handleHeaderFooterSave,
    handleBodyClick,
    handleRemoveHeaderFooter,
  };
};
