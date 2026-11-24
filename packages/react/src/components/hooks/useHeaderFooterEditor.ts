import { useCallback, useMemo, useState } from "react";

/**
 * Hook encapsulating header/footer editing state, content resolution, and
 * the canonical operations that create and remove a header or footer.
 */
import {
  createCanonicalHeaderFooterOperation,
  removeCanonicalHeaderFooterOperations,
} from "@stll/folio-core/canonical/canonicalOperations";
import type { FolioEditor } from "@stll/folio-core/controller/folioEditor";
import type {
  HeaderFooter,
  SectionProperties,
} from "@stll/folio-core/types/document";
import {
  type DeepReadonly,
  readModel,
  type ReadonlyDocument,
} from "@stll/folio-core/types/documentModel";
import {
  pickActiveHeaderFooterRId,
  resolveEffectiveSectionProperties,
  resolveHeaderFooterContent,
} from "@stll/folio-core/utils/headerFooter";

import type { UseHistoryReturn } from "../../hooks/useHistory";

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

type UseHeaderFooterEditorParams = {
  history: UseHistoryReturn<ReadonlyDocument | null>;
  getCanonicalApi?: () => Pick<
    FolioEditor,
    "ensureView" | "getCanonicalDocument" | "applyCanonicalOperations"
  > | null;
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

  /**
   * Section properties with titlePg merged from inline sections: the model's
   * own when nothing is merged, so read them, never write them.
   */
  effectiveSectionProperties: DeepReadonly<SectionProperties> | undefined;

  /** Open the inline HF editor on double-click; creates an empty HF if needed */
  handleHeaderFooterDoubleClick: (
    position: "header" | "footer",
    pageNumber?: number,
  ) => void;
  /** Close the HF editor; the session already holds its edits. */
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
  getCanonicalApi,
}: UseHeaderFooterEditorParams): UseHeaderFooterEditorReturn => {
  // -------------------------------------------------------------------------
  // State
  // -------------------------------------------------------------------------

  const [hfEditPosition, setHfEditPosition] = useState<
    "header" | "footer" | null
  >(null);
  const [hfEditIsFirstPage, setHfEditIsFirstPage] = useState(false);

  // -------------------------------------------------------------------------
  // Resolved header/footer content
  // -------------------------------------------------------------------------

  const rawResolution = useMemo(
    () =>
      resolveHeaderFooterContent(
        history.state ? readModel(history.state).package : undefined,
      ),
    [history.state],
  );
  // `history.state` is transiently null while the document buffer is parsed and
  // during relayout. Falling through to the empty resolution on those renders
  // blanks the active H/F rIds, which breaks edit focus (getActiveView returns
  // null, so typing lands in the body) and flickers the rendered chrome. Hold
  // the last resolved structure and reuse it until a real document state
  // arrives again.
  const [lastResolution, setLastResolution] = useState(rawResolution);
  if (
    history.state?.package &&
    Object.entries(rawResolution).some(
      ([key, value]) => Reflect.get(lastResolution, key) !== value,
    )
  ) {
    setLastResolution(rawResolution);
  }
  const resolution = history.state?.package ? rawResolution : lastResolution;
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
    () =>
      resolveEffectiveSectionProperties(
        history.state?.package.document,
        hasTitlePg,
      ),
    [history.state, hasTitlePg],
  );

  // -------------------------------------------------------------------------
  // Callbacks
  // -------------------------------------------------------------------------

  const handleHeaderFooterDoubleClick = useCallback(
    (position: "header" | "footer", pageNumber?: number) => {
      const isFirstPage = hasTitlePg && (pageNumber ?? 1) === 1;
      let hf = position === "header" ? headerContent : footerContent;
      if (isFirstPage) {
        hf =
          position === "header"
            ? firstPageHeaderContent
            : firstPageFooterContent;
      }
      setHfEditIsFirstPage(isFirstPage);
      if (hf) {
        setHfEditPosition(position);
        return;
      }

      // Create an empty header/footer for docs that don't have one yet.
      const api = getCanonicalApi?.();
      api?.ensureView();
      const canonical = api?.getCanonicalDocument();
      if (!canonical || !api) return;
      const operation = createCanonicalHeaderFooterOperation({
        document: canonical,
        position,
        referenceType: isFirstPage ? "first" : "default",
      });
      if (api.applyCanonicalOperations([operation]))
        setHfEditPosition(position);
    },
    [
      headerContent,
      footerContent,
      firstPageHeaderContent,
      firstPageFooterContent,
      hasTitlePg,
      getCanonicalApi,
    ],
  );

  // The session holds every story's edits; closing the editor saves nothing.
  const handleHeaderFooterSave = useCallback(() => {
    setHfEditPosition(null);
  }, []);

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
    const activeRId = pickActiveHeaderFooterRId(
      resolution,
      hfEditPosition,
      hfEditIsFirstPage,
    );
    const api = getCanonicalApi?.();
    const canonical = api?.getCanonicalDocument();
    if (api && canonical && activeRId) {
      api.applyCanonicalOperations(
        removeCanonicalHeaderFooterOperations({
          document: canonical,
          position: hfEditPosition,
          rId: activeRId,
        }),
      );
    }
    setHfEditPosition(null);
  }, [hfEditPosition, hfEditIsFirstPage, resolution, history, getCanonicalApi]);

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
