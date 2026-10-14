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
import { panic } from "better-result";
import type { Node as PMNode } from "prosemirror-model";

import { FindReplaceManager } from "@stll/folio-core/managers/FindReplaceManager";
import { replaceEligibility } from "@stll/folio-core/utils/findReplace";
import type { ReplaceAllSummary } from "@stll/folio-core/utils/findReplace";
import {
  createFindSearchKey,
  createFindQueryKey,
  sameFindSearchKey,
  sameFindBodyVersion,
} from "@stll/folio-core/utils/findSearchKey";
import type { Document } from "@stll/folio-core/types/document";
import type { DisplayMode } from "@stll/folio-core/managers/EditorModeManager";
import {
  createFindMatchView,
  getRenderedSearchStories,
} from "@stll/folio-core/utils/findMatchView";
import type { Layout } from "@stll/folio-core/layout-engine/types";
import { findInDocument } from "../dialogs/findReplaceUtils";
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
  /** Ref to the document search surface */
  containerRef: React.RefObject<HTMLDivElement | null>;
  /** Callback to push a new document state */
  handleDocumentChange: (newDoc: Document) => void;
  /** Dialog state manager from useFindReplace (dialogs) */
  findReplace: FindReplaceStateReturn;
  /** Activate the thread containing a comment match */
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
  currentSearch: { queryKey: string; result: FindResult } | null;
  /** Clear results and navigation when the document or query is reset. */
  resetFindResult: () => void;
  /** Execute a find operation */
  handleFind: (searchText: string, options: FindOptions) => FindResult | null;
  /** Publish query edits before the debounced search completes. */
  handleQueryChange: (searchText: string, options: FindOptions) => void;
  /** Navigate to the next match */
  handleFindNext: () => FindMatch | null;
  /** Navigate to the previous match */
  handleFindPrevious: () => FindMatch | null;
  /** Replace the current match */
  handleReplace: (replaceText: string) => boolean;
  /** Replace all matches */
  handleReplaceAll: (
    searchText: string,
    replaceText: string,
    options: FindOptions,
  ) => ReplaceAllSummary;
  getFindResult: () => FindResult | null;
  refreshFind: () => void;
};

type PublishedFindResult =
  | { status: "cleared" }
  | { status: "ready"; queryKey: string; result: FindResult };

type ReplacementGate =
  | { status: "ready" }
  | { status: "awaiting-document"; source: ReturnType<typeof createFindSearchKey> };

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
  const [publishedResult, setPublishedResult] = useState<PublishedFindResult>({
    status: "cleared",
  });
  const { setMatches, goToMatch } = findReplace;
  const matchView = useRef<ReturnType<typeof createFindMatchView> | null>(null);
  const query = useRef<{ text: string; options: FindOptions } | null>(null);
  const searchSource = useRef<
    (ReturnType<typeof createFindSearchKey> & { intentKey: string }) | null
  >(null);
  const replacementGate = useRef<ReplacementGate>({ status: "ready" });
  const resultCallback = useRef(onFindResultChange);
  useEffect(() => {
    resultCallback.current = onFindResultChange;
  }, [onFindResultChange]);
  const publish = useCallback((result: FindResult | null) => {
    if (result) {
      const source = searchSource.current;
      if (!source) panic("Find result published without a search key");
      setPublishedResult({ status: "ready", queryKey: source.intentKey, result });
    } else setPublishedResult({ status: "cleared" });
    resultCallback.current?.(result);
  }, []);

  useEffect(() => () => matchView.current?.destroy(), []);

  const clearFindResult = useCallback(() => {
    manager.clear();
    searchSource.current = null;
    matchView.current?.destroy();
    matchView.current = null;
    publish(null);
    setMatches([], 0);
  }, [manager, setMatches, publish]);
  const resetFindResult = useCallback(() => {
    query.current = null;
    clearFindResult();
  }, [clearFindResult]);

  const readDocumentState = useCallback(
    () => getDocumentState?.() ?? documentState ?? null,
    [getDocumentState, documentState],
  );

  const revealMatch = useCallback(
    (match: FindMatch) => {
      const index = manager.getResult()?.currentIndex ?? 0;
      matchView.current?.navigate(index, () => selectMatch?.(match));
    },
    [selectMatch, manager],
  );

  const handleQueryChange = useCallback((text: string, options: FindOptions) => {
    query.current = { text, options };
  }, []);

  const handleFind = useCallback(
    (searchText: string, options: FindOptions): FindResult | null => {
      const currentDocument = readDocumentState();
      if (!currentDocument || !searchText.trim()) {
        resetFindResult();
        return null;
      }

      const bodyDoc = getBodyDoc?.();
      const effectiveOptions = {
        ...options,
        includeComments: true,
        ...(displayMode && { displayMode }),
      };
      const context = getLayout
        ? { renderedStories: getRenderedSearchStories(getLayout()) }
        : undefined;
      const nextSource = createFindSearchKey({
        query: searchText,
        options: effectiveOptions,
        document: currentDocument,
        ...(bodyDoc && { bodyVersion: bodyDoc }),
        ...(context && { context }),
      });
      const refresh = searchSource.current?.query === nextSource.query;
      const currentIndex = refresh ? (manager.getResult()?.currentIndex ?? 0) : 0;
      query.current = { text: searchText, options };
      searchSource.current = { ...nextSource, intentKey: createFindQueryKey(searchText, options) };
      const matches = findInDocument(currentDocument, searchText, effectiveOptions, context);
      const result = manager.setMatches(matches, currentIndex);
      matchView.current?.destroy();
      const container = containerRef.current;
      matchView.current = container
        ? createFindMatchView({
            container,
            document: currentDocument,
            bodyDoc,
            matches,
            layout: getLayout?.(),
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

  const refreshChangedSource = useCallback(
    (currentDocument: Document): boolean => {
      const active = query.current;
      if (!active) return false;
      const bodyDoc = getBodyDoc?.();
      const context = getLayout
        ? { renderedStories: getRenderedSearchStories(getLayout()) }
        : undefined;
      const currentSource = createFindSearchKey({
        query: active.text,
        options: { ...active.options, includeComments: true, ...(displayMode && { displayMode }) },
        document: currentDocument,
        ...(bodyDoc && { bodyVersion: bodyDoc }),
        ...(context && { context }),
      });
      switch (replacementGate.current.status) {
        case "awaiting-document":
          if (sameFindBodyVersion(replacementGate.current.source, currentSource)) return true;
          replacementGate.current = { status: "ready" };
          break;
        case "ready":
          break;
        default: {
          const exhaustive: never = replacementGate.current;
          return exhaustive;
        }
      }
      if (sameFindSearchKey(searchSource.current, currentSource)) return false;
      handleFind(active.text, active.options);
      return true;
    },
    [getBodyDoc, getLayout, handleFind, displayMode],
  );

  const acceptReplacement = useCallback(
    (nextDocument: Document) => {
      const source = searchSource.current;
      if (!source) panic("Replacement accepted without a search key");
      replacementGate.current = { status: "awaiting-document", source };
      clearFindResult();
      handleDocumentChange(nextDocument);
    },
    [handleDocumentChange, clearFindResult],
  );

  const handleReplace = useCallback(
    (replaceText: string): boolean => {
      if (readOnly) return false;
      const currentDocument = readDocumentState();
      if (!currentDocument) {
        return false;
      }
      if (refreshChangedSource(currentDocument)) return false;
      const result = manager.getResult();
      const match = result?.matches.at(result.currentIndex);
      if (!match || replaceEligibility(match).status === "blocked") return false;
      if (match.text === replaceText) return true;

      const newDoc = manager.replaceCurrent(currentDocument, replaceText);
      if (!newDoc) {
        return false;
      }

      acceptReplacement(newDoc);
      return true;
    },
    [readDocumentState, manager, acceptReplacement, readOnly, refreshChangedSource],
  );

  const handleReplaceAll = useCallback(
    (searchText: string, replaceText: string, options: FindOptions): ReplaceAllSummary => {
      if (readOnly) return { replacedCount: 0, skippedCount: manager.getResult()?.totalCount ?? 0 };
      const currentDocument = readDocumentState();
      if (!currentDocument || !searchText.trim()) {
        return { replacedCount: 0, skippedCount: 0 };
      }
      const previousCount = manager.getResult()?.matches.length ?? 0;
      handleQueryChange(searchText, options);
      if (refreshChangedSource(currentDocument))
        return { replacedCount: 0, skippedCount: previousCount };

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
      let unchangedCount = 0;
      const mutationMatches = matches.filter((match) => {
        if (replaceEligibility(match).status === "allowed" && match.text === replaceText) {
          unchangedCount++;
          return false;
        }
        return true;
      });
      if (mutationMatches.length === 0) return { replacedCount: unchangedCount, skippedCount: 0 };
      const outcome = manager.replaceAll(currentDocument, mutationMatches, replaceText);
      if (!outcome) {
        return { replacedCount: 0, skippedCount: 0 };
      }

      if (outcome.replacedCount > 0) {
        acceptReplacement(outcome.document);
      }
      return {
        replacedCount: outcome.replacedCount + unchangedCount,
        skippedCount: outcome.skippedCount,
      };
    },
    [
      readDocumentState,
      manager,
      acceptReplacement,
      readOnly,
      displayMode,
      getLayout,
      refreshChangedSource,
      handleQueryChange,
    ],
  );

  return {
    currentResult: publishedResult.status === "ready" ? publishedResult.result : null,
    currentSearch:
      publishedResult.status === "ready"
        ? { queryKey: publishedResult.queryKey, result: publishedResult.result }
        : null,
    resetFindResult,
    handleFind,
    handleQueryChange,
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
