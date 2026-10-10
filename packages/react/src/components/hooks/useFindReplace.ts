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

import { useCallback, useMemo, useState } from "react";

import { FindReplaceManager } from "@stll/folio-core/managers/FindReplaceManager";
import type { Document } from "@stll/folio-core/types/document";
import { findInDocument, scrollToMatch } from "../dialogs/findReplaceUtils";
import type { FindMatch, FindOptions, FindResult } from "../dialogs/findReplaceUtils";
import type { UseFindReplaceReturn as FindReplaceStateReturn } from "../dialogs/useFindReplace";

// ============================================================================
// TYPES
// ============================================================================

type UseFindReplaceParams = {
  /** Current document state fallback for existing mounted callers during HMR */
  documentState?: Document | null;
  /** Returns the current live document state from the editor */
  getDocumentState?: () => Document | null;
  /** Ref to the scrollable container for scrollToMatch */
  containerRef: React.RefObject<HTMLDivElement | null>;
  /** Callback to push a new document state */
  handleDocumentChange: (newDoc: Document) => void;
  /** Dialog state manager from useFindReplace (dialogs) */
  findReplace: FindReplaceStateReturn;
  /** Select and reveal a match in the live editor */
  selectMatch?: (match: FindMatch) => boolean;
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
  containerRef,
  handleDocumentChange,
  findReplace,
  selectMatch,
}: UseFindReplaceParams): UseFindReplaceReturn {
  const manager = useMemo(() => new FindReplaceManager<FindMatch>(), []);
  const [currentResult, setCurrentResult] = useState<FindResult | null>(null);
  const { setMatches, goToMatch } = findReplace;

  const resetFindResult = useCallback(() => {
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

      const matches = findInDocument(currentDocument, searchText, options);
      const result = manager.setMatches(matches);
      setCurrentResult(result);
      setMatches(matches, 0);

      if (matches.length > 0) {
        // SAFETY: length > 0 guarantees index 0 exists
        revealMatch(matches[0]!);
      }

      return result;
    },
    [readDocumentState, manager, setMatches, revealMatch, resetFindResult],
  );

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

  return {
    currentResult,
    resetFindResult,
    handleFind,
    handleFindNext,
    handleFindPrevious,
    handleReplace,
    handleReplaceAll,
  };
}
