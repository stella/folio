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
import type { Node as PMNode } from "prosemirror-model";

import { FindReplaceManager } from "@stll/folio-core/managers/FindReplaceManager";
import type { Document } from "@stll/folio-core/types/document";
import type { DisplayMode } from "@stll/folio-core/managers/EditorModeManager";
import {
  createFindMatchView,
  getRenderedSearchStories,
} from "@stll/folio-core/utils/findMatchView";
import type { Layout } from "@stll/folio-core/layout-engine/types";
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
  getBodyDoc?: () => PMNode | undefined;
  getLayout?: () => Layout | null;
  displayMode?: DisplayMode;
  onFindResultChange?: ((result: FindResult | null) => void) | undefined;
  readOnly?: boolean;
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
  getFindResult: () => FindResult | null;
  refreshFind: () => void;
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
  getBodyDoc,
  getLayout,
  displayMode,
  onFindResultChange,
  readOnly,
}: UseFindReplaceParams): UseFindReplaceReturn {
  const manager = useMemo(() => new FindReplaceManager<FindMatch>(), []);
  const [currentResult, setCurrentResult] = useState<FindResult | null>(null);
  const { setMatches, goToMatch } = findReplace;
  const matchView = useRef<ReturnType<typeof createFindMatchView> | null>(null);
  const query = useRef<{ text: string; options: FindOptions } | null>(null);
  const resultCallback = useRef(onFindResultChange);
  useEffect(() => {
    resultCallback.current = onFindResultChange;
  }, [onFindResultChange]);
  const publish = useCallback((result: FindResult | null) => {
    setCurrentResult(result);
    resultCallback.current?.(result);
  }, []);

  useEffect(() => () => matchView.current?.destroy(), []);

  const resetFindResult = useCallback(() => {
    manager.clear();
    query.current = null;
    matchView.current?.destroy();
    matchView.current = null;
    publish(null);
    setMatches([], 0);
  }, [manager, setMatches, publish]);

  const readDocumentState = useCallback(
    () => getDocumentState?.() ?? documentState ?? null,
    [getDocumentState, documentState],
  );

  const revealMatch = useCallback(
    (match: FindMatch) => {
      const index = manager.getResult()?.currentIndex ?? 0;
      matchView.current?.paint(index);
      if (matchView.current?.reveal(index)) return;
      if (selectMatch?.(match)) return;
      if (
        (match.story === undefined || match.story === "main") &&
        match.textBoxIndex === undefined &&
        match.commentId === undefined &&
        containerRef.current
      ) {
        scrollToMatch(containerRef.current, match);
      }
    },
    [selectMatch, containerRef, manager],
  );

  const handleFind = useCallback(
    (searchText: string, options: FindOptions): FindResult | null => {
      const currentDocument = readDocumentState();
      if (!currentDocument || !searchText.trim()) {
        resetFindResult();
        return null;
      }

      const previous = query.current;
      const refresh =
        previous?.text === searchText &&
        previous.options.matchCase === options.matchCase &&
        previous.options.matchWholeWord === options.matchWholeWord &&
        previous.options.useRegex === options.useRegex;
      const currentIndex = refresh ? (manager.getResult()?.currentIndex ?? 0) : 0;
      query.current = { text: searchText, options };
      const matches = findInDocument(
        currentDocument,
        searchText,
        { ...options, includeComments: true, ...(displayMode && { displayMode }) },
        getLayout ? { renderedStories: getRenderedSearchStories(getLayout()) } : undefined,
      );
      const result = manager.setMatches(matches, currentIndex);
      matchView.current?.destroy();
      const container = containerRef.current;
      matchView.current =
        container && getBodyDoc
          ? createFindMatchView({
              container,
              document: currentDocument,
              bodyDoc: getBodyDoc(),
              matches,
            })
          : null;
      publish(result);
      setMatches(matches, result.currentIndex);
      matchView.current?.paint(result.currentIndex);

      if (matches.length > 0 && !refresh) {
        // SAFETY: length > 0 guarantees index 0 exists
        revealMatch(matches[0]!);
      }

      return result;
    },
    [
      readDocumentState,
      manager,
      setMatches,
      revealMatch,
      resetFindResult,
      displayMode,
      containerRef,
      getBodyDoc,
      getLayout,
      publish,
    ],
  );

  useEffect(() => {
    if (!query.current) return;
    const timer = setTimeout(() => {
      const active = query.current;
      if (active) handleFind(active.text, active.options);
    }, 200);
    return () => clearTimeout(timer);
  }, [handleFind]);

  const handleFindNext = useCallback((): FindMatch | null => {
    const stepped = manager.navigate("next");
    if (!stepped) {
      return null;
    }
    publish(manager.getResult());
    goToMatch(stepped.index);
    revealMatch(stepped.match);
    return stepped.match;
  }, [manager, goToMatch, revealMatch, publish]);

  const handleFindPrevious = useCallback((): FindMatch | null => {
    const stepped = manager.navigate("previous");
    if (!stepped) {
      return null;
    }
    publish(manager.getResult());
    goToMatch(stepped.index);
    revealMatch(stepped.match);
    return stepped.match;
  }, [manager, goToMatch, revealMatch, publish]);

  const handleReplace = useCallback(
    (replaceText: string): boolean => {
      if (readOnly) return false;
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
    [readDocumentState, manager, handleDocumentChange, readOnly],
  );

  const handleReplaceAll = useCallback(
    (searchText: string, replaceText: string, options: FindOptions): number => {
      if (readOnly) return 0;
      const currentDocument = readDocumentState();
      if (!currentDocument || !searchText.trim()) {
        return 0;
      }

      const matches = findInDocument(
        currentDocument,
        searchText,
        {
          ...options,
          includeComments: true,
          ...(displayMode && { displayMode }),
        },
        getLayout ? { renderedStories: getRenderedSearchStories(getLayout()) } : undefined,
      );
      const outcome = manager.replaceAll(currentDocument, matches, replaceText);
      if (!outcome) {
        return 0;
      }

      handleDocumentChange(outcome.document);
      resetFindResult();
      return outcome.replacedCount;
    },
    [
      readDocumentState,
      manager,
      handleDocumentChange,
      resetFindResult,
      readOnly,
      displayMode,
      getLayout,
    ],
  );

  return {
    currentResult,
    resetFindResult,
    handleFind,
    handleFindNext,
    handleFindPrevious,
    handleReplace,
    handleReplaceAll,
    getFindResult: () => manager.getResult(),
    refreshFind: () => {
      const active = query.current;
      if (active) handleFind(active.text, active.options);
    },
  };
}
