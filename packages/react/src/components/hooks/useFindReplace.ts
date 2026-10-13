/**
 * useFindReplace Hook
 *
 * Thin React binding around the framework-agnostic FindReplaceManager. The
 * manager owns the active find result (matches + cursor) and runs the
 * document-level replace operations; this hook keeps the React glue: the
 * `currentResult` the dialog reads, the live-document accessor, scroll/select
 * side effects, and the dialog-state wiring.
 *
 * The dialog visibility/search-text state itself is managed by `useFindReplace`
 * from `dialogs/useFindReplace.ts`; this hook layers the document-aware
 * operations (find, replace, scroll-to-match) on top.
 */

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import type { EditorState } from "prosemirror-state";

import { FindReplaceManager } from "@stll/folio-core/managers/FindReplaceManager";
import { getReviewViewReadOnlyReason } from "@stll/folio-core/managers/EditorModeManager";
import type {
  DisplayMode,
  EditorReadOnlyReason,
} from "@stll/folio-core/managers/EditorModeManager";
import { findInMarkupView } from "@stll/folio-core/prosemirror/findReplaceSelection";
import type { ProseMirrorFindMatch } from "@stll/folio-core/prosemirror/findReplaceSelection";
import type { Document } from "@stll/folio-core/types/document";
import { findInDocument, scrollToMatch } from "../dialogs/findReplaceUtils";
import type { FindMatch, FindOptions, FindResult } from "../dialogs/findReplaceUtils";
import type { UseFindReplaceReturn as FindReplaceStateReturn } from "../dialogs/useFindReplace";

const toEditorFindMatch = ({ from, to, ...match }: ProseMirrorFindMatch) => ({
  ...match,
  contentIndex: 0,
  authoredRange: { from, to },
});

// ============================================================================
// TYPES
// ============================================================================

type UseFindReplaceParams = {
  /** Current document state fallback for existing mounted callers during HMR */
  documentState?: Document | null;
  /** Returns the current live document state from the editor */
  getDocumentState?: () => Document | null;
  /** Live editor state for searching the displayed review view. */
  getEditorState?: () => EditorState | null;
  markupView?: DisplayMode;
  /** Ref to the scrollable container for scrollToMatch */
  containerRef: React.RefObject<HTMLDivElement | null>;
  /** Callback to push a new document state */
  handleDocumentChange: (newDoc: Document) => void;
  /** Dialog state manager from useFindReplace (dialogs) */
  findReplace: FindReplaceStateReturn;
  /** Select and reveal a match in the live editor */
  selectMatch?: (match: FindMatch) => boolean;
  readOnlyReason?: EditorReadOnlyReason | null;
  onReadonlyEditAttempt?: (reason: EditorReadOnlyReason) => void;
};

export type UseFindReplaceReturn = {
  /** Reactive current find result for FindReplaceDialog. */
  currentResult: FindResult | null;
  /** Clear results and navigation when the document or query is reset. */
  resetFindResult: () => void;
  /** Execute a find operation */
  handleFind: (searchText: string, options: FindOptions) => FindResult | null;
  /** Navigate to the next match */
  handleFindNext: () => FindMatch | null;
  /** Navigate to the previous match */
  handleFindPrevious: () => FindMatch | null;
  /** Replace the current match */
  handleReplace: (replaceText: string) => boolean;
  /** Replace all matches */
  handleReplaceAll: (searchText: string, replaceText: string, options: FindOptions) => number;
};

// ============================================================================
// HOOK
// ============================================================================

export function useFindReplace({
  documentState,
  getDocumentState,
  getEditorState,
  markupView: markupViewProp,
  containerRef,
  handleDocumentChange,
  findReplace,
  selectMatch,
  readOnlyReason,
  onReadonlyEditAttempt,
}: UseFindReplaceParams): UseFindReplaceReturn {
  const markupView = markupViewProp ?? "all-markup";
  const replacementRefusal = getReviewViewReadOnlyReason(markupView) ?? readOnlyReason ?? null;
  const manager = useMemo(() => new FindReplaceManager<FindMatch>(), []);
  const [currentResult, setCurrentResult] = useState<FindResult | null>(null);
  const activeSearch = useRef<{ searchText: string; options: FindOptions } | null>(null);
  const { setMatches, goToMatch } = findReplace;

  const resetFindResult = useCallback(() => {
    activeSearch.current = null;
    manager.clear();
    setCurrentResult(null);
    setMatches([], 0);
  }, [manager, setMatches]);

  const readDocumentState = useCallback(
    () => getDocumentState?.() ?? documentState ?? null,
    [getDocumentState, documentState],
  );

  const revealMatch = useCallback(
    (match: FindMatch) => {
      if (!selectMatch?.(match) && containerRef.current) {
        scrollToMatch(containerRef.current, match);
      }
    },
    [selectMatch, containerRef],
  );

  const handleFind = useCallback(
    (searchText: string, options: FindOptions): FindResult | null => {
      const currentDocument = readDocumentState();
      if (!currentDocument || !searchText.trim()) {
        resetFindResult();
        return null;
      }

      activeSearch.current = { searchText, options };
      const editorState = getEditorState?.();
      const matches =
        editorState && markupView !== "all-markup"
          ? findInMarkupView({ state: editorState, markupView, searchText, options }).map(
              toEditorFindMatch,
            )
          : findInDocument(currentDocument, searchText, options);
      const result = manager.setMatches(matches);
      setCurrentResult(result);
      setMatches(matches, 0);

      if (matches.length > 0) {
        // SAFETY: length > 0 guarantees index 0 exists
        revealMatch(matches[0]!);
      }

      return result;
    },
    [
      readDocumentState,
      getEditorState,
      markupView,
      manager,
      setMatches,
      revealMatch,
      resetFindResult,
    ],
  );

  const previousMarkupView = useRef(markupView);
  useEffect(() => {
    if (previousMarkupView.current === markupView) return;
    previousMarkupView.current = markupView;
    const search = activeSearch.current;
    if (search) handleFind(search.searchText, search.options);
  }, [markupView, handleFind]);

  const handleFindNext = useCallback((): FindMatch | null => {
    const stepped = manager.navigate("next");
    if (!stepped) {
      return null;
    }
    setCurrentResult(manager.getResult());
    goToMatch(stepped.index);
    revealMatch(stepped.match);
    return stepped.match;
  }, [manager, goToMatch, revealMatch]);

  const handleFindPrevious = useCallback((): FindMatch | null => {
    const stepped = manager.navigate("previous");
    if (!stepped) {
      return null;
    }
    setCurrentResult(manager.getResult());
    goToMatch(stepped.index);
    revealMatch(stepped.match);
    return stepped.match;
  }, [manager, goToMatch, revealMatch]);

  const handleReplace = useCallback(
    (replaceText: string): boolean => {
      const currentDocument = readDocumentState();
      if (!currentDocument) {
        return false;
      }

      const newDoc = manager.replaceCurrent(currentDocument, replaceText);
      if (!newDoc) {
        return false;
      }

      handleDocumentChange(newDoc);
      return true;
    },
    [readDocumentState, manager, handleDocumentChange],
  );

  const handleReplaceAll = useCallback(
    (searchText: string, replaceText: string, options: FindOptions): number => {
      const currentDocument = readDocumentState();
      if (!currentDocument || !searchText.trim()) {
        return 0;
      }

      const matches = findInDocument(currentDocument, searchText, options);
      const outcome = manager.replaceAll(currentDocument, matches, replaceText);
      if (!outcome) {
        return 0;
      }

      handleDocumentChange(outcome.document);
      resetFindResult();
      return outcome.replacedCount;
    },
    [readDocumentState, manager, handleDocumentChange, resetFindResult],
  );

  const refuseReplace = useCallback(() => {
    if (replacementRefusal) onReadonlyEditAttempt?.(replacementRefusal);
    return false;
  }, [replacementRefusal, onReadonlyEditAttempt]);
  const refuseReplaceAll = useCallback(() => {
    refuseReplace();
    return 0;
  }, [refuseReplace]);

  return {
    currentResult,
    resetFindResult,
    handleFind,
    handleFindNext,
    handleFindPrevious,
    handleReplace: replacementRefusal ? refuseReplace : handleReplace,
    handleReplaceAll: replacementRefusal ? refuseReplaceAll : handleReplaceAll,
  };
}
