/**
 * Find and Replace Dialog Component
 *
 * Modal dialog for searching and replacing text in the document.
 * Supports find, find next/previous, replace, and replace all operations.
 *
 * Logic and utilities are in separate files:
 * - findReplaceUtils.ts — Pure search/replace functions and types
 * - useFindReplace.ts   — React hook for dialog state management
 */

import React, { useState, useEffect, useLayoutEffect, useCallback, useRef } from "react";
import type { CSSProperties, KeyboardEvent, ChangeEvent } from "react";

import { ChevronDownIcon, ChevronUpIcon, SearchIcon, XIcon } from "lucide-react";
import { useTranslations } from "use-intl";

import { useFolioUI } from "../../ui/folio-ui";
import { getFindReplaceOverlayStyle } from "./findReplaceDialogLayout";
import { getFindEnterAction } from "./findReplaceInteraction";
import type { FindOptions, FindResult, FindMatch } from "./findReplaceUtils";

// Re-export types and utilities so existing imports still work
export type { FindMatch, FindOptions, FindResult, HighlightOptions } from "./findReplaceUtils";
export {
  createDefaultFindOptions,
  findAllMatches,
  escapeRegexString,
  createSearchPattern,
  replaceAllInContent,
  replaceFirstInContent,
  getMatchCountText,
  isEmptySearch,
  getDefaultHighlightOptions,
  findInDocument,
  findInParagraph,
  scrollToMatch,
} from "./findReplaceUtils";

export type { FindReplaceOptions, FindReplaceState, UseFindReplaceReturn } from "./useFindReplace";
export { useFindReplace } from "./useFindReplace";

// ============================================================================
// PROPS
// ============================================================================

/**
 * Props for the FindReplaceDialog component
 */
export type FindReplaceDialogProps = {
  /** Whether the dialog is open */
  isOpen: boolean;
  /** Callback when dialog is closed */
  onClose: () => void;
  /** Callback when searching for text */
  onFind: (searchText: string, options: FindOptions) => FindResult | null;
  /** Callback when navigating to next match */
  onFindNext: () => FindMatch | null;
  /** Callback when navigating to previous match */
  onFindPrevious: () => FindMatch | null;
  /** Callback when replacing current match */
  onReplace: (replaceText: string) => boolean;
  /** Callback when replacing all matches */
  onReplaceAll: (searchText: string, replaceText: string, options: FindOptions) => number;
  /** Callback to highlight matches in document */
  onHighlightMatches?: (matches: FindMatch[]) => void;
  /** Callback to clear highlights */
  onClearHighlights?: () => void;
  /** Initial search text (e.g., from selected text) */
  initialSearchText?: string;
  /** Whether to start in replace mode */
  replaceMode?: boolean;
  /** Current match result (from external state) */
  currentResult?: FindResult | null;
  /** Additional CSS class */
  className?: string;
  /** Additional inline styles */
  style?: CSSProperties;
};

// ============================================================================
// MAIN COMPONENT
// ============================================================================

/**
 * FindReplaceDialog component - Modal for finding and replacing text
 */
export function FindReplaceDialog(props: FindReplaceDialogProps): React.ReactElement | null {
  // Search preferences survive closing; each opening starts a fresh query.
  const [matchCase, setMatchCase] = useState(false);
  const [matchWholeWord, setMatchWholeWord] = useState(false);
  const { isOpen, onClearHighlights } = props;
  useEffect(() => {
    if (!isOpen) {
      onClearHighlights?.();
    }
  }, [isOpen, onClearHighlights]);
  if (!isOpen) {
    return null;
  }
  return (
    <FindReplaceDialogForm
      key={props.initialSearchText}
      {...props}
      matchCase={matchCase}
      matchWholeWord={matchWholeWord}
      onMatchCaseChange={setMatchCase}
      onMatchWholeWordChange={setMatchWholeWord}
    />
  );
}

type FindReplaceDialogFormProps = FindReplaceDialogProps & {
  matchCase: boolean;
  matchWholeWord: boolean;
  onMatchCaseChange: (value: boolean) => void;
  onMatchWholeWordChange: (value: boolean) => void;
};

function FindReplaceDialogForm({
  isOpen,
  onClose,
  onFind,
  onFindNext,
  onFindPrevious,
  onHighlightMatches,
  onClearHighlights,
  initialSearchText = "",
  currentResult,
  className,
  style,
  matchCase,
  matchWholeWord,
  onMatchCaseChange,
  onMatchWholeWordChange,
}: FindReplaceDialogFormProps): React.ReactElement | null {
  const id = React.useId();
  const t = useTranslations("folio");
  const { Button, Input, Checkbox } = useFolioUI();
  // State
  const [searchText, setSearchText] = useState(initialSearchText);
  const [localResult, setResult] = useState<FindResult | null>(currentResult ?? null);
  const [previousCurrentResult, setPreviousCurrentResult] = useState(currentResult);
  if (currentResult !== previousCurrentResult) {
    setPreviousCurrentResult(currentResult);
    if (currentResult !== undefined) {
      setResult(currentResult);
    }
  }
  // Host results update the local cursor only when a new result arrives. Local
  // query/options edits can invalidate it while the host still holds the old one.
  const result = searchText.trim() ? localResult : null;

  // Refs
  const searchInputRef = useRef<HTMLInputElement>(null);
  useEffect(() => {
    const timeout = setTimeout(() => {
      searchInputRef.current?.focus();
      searchInputRef.current?.select();
    }, 100);
    return () => clearTimeout(timeout);
  }, []);

  const performSearch = useCallback(
    (text = searchText, options = { matchCase, matchWholeWord }) => {
      if (!text.trim()) {
        setResult(null);
        onClearHighlights?.();
        return;
      }
      const searchResult = onFind(text, options);
      setResult(searchResult);
      if (searchResult?.matches && onHighlightMatches) {
        onHighlightMatches(searchResult.matches);
      } else {
        onClearHighlights?.();
      }
    },
    [searchText, matchCase, matchWholeWord, onFind, onHighlightMatches, onClearHighlights],
  );

  // Keep callback updates out of the query schedule. A search can update the
  // host's callbacks without scheduling another search for the same query.
  const performSearchRef = useRef(performSearch);
  const clearHighlightsRef = useRef(onClearHighlights);
  useLayoutEffect(() => {
    performSearchRef.current = performSearch;
    clearHighlightsRef.current = onClearHighlights;
  }, [performSearch, onClearHighlights]);

  useEffect(() => {
    if (!searchText.trim()) {
      clearHighlightsRef.current?.();
      return;
    }
    const timeout = setTimeout(
      () => performSearchRef.current(searchText, { matchCase, matchWholeWord }),
      120,
    );
    return () => clearTimeout(timeout);
  }, [searchText, matchCase, matchWholeWord]);

  const handleMatchCaseChange = useCallback(
    (value: boolean) => {
      setResult(null);
      onMatchCaseChange(value);
    },
    [onMatchCaseChange],
  );
  const handleMatchWholeWordChange = useCallback(
    (value: boolean) => {
      setResult(null);
      onMatchWholeWordChange(value);
    },
    [onMatchWholeWordChange],
  );

  const handleSearchChange = useCallback((e: ChangeEvent<HTMLInputElement>) => {
    setSearchText(e.target.value);
    setResult(null);
  }, []);

  const handleFindNext = useCallback(() => {
    if (!searchText.trim()) {
      performSearch();
      return;
    }

    if (!result || result.totalCount === 0) {
      performSearch();
      return;
    }

    const match = onFindNext();
    if (match) {
      const newIndex = (result.currentIndex + 1) % result.totalCount;
      setResult({
        ...result,
        currentIndex: newIndex,
      });
    }
  }, [searchText, result, performSearch, onFindNext]);

  const handleFindPrevious = useCallback(() => {
    if (!searchText.trim()) {
      performSearch();
      return;
    }

    if (!result || result.totalCount === 0) {
      performSearch();
      return;
    }

    const match = onFindPrevious();
    if (match) {
      const newIndex = result.currentIndex === 0 ? result.totalCount - 1 : result.currentIndex - 1;
      setResult({
        ...result,
        currentIndex: newIndex,
      });
    }
  }, [searchText, result, performSearch, onFindPrevious]);

  const handleSearchKeyDown = useCallback(
    (e: KeyboardEvent<HTMLInputElement>) => {
      if (e.key === "Enter") {
        e.preventDefault();
        const action = getFindEnterAction({
          searchText,
          result,
          shiftKey: e.shiftKey,
        });
        if (action === "search") {
          performSearch();
          return;
        }
        if (action === "previous") {
          handleFindPrevious();
          return;
        }
        handleFindNext();
      } else if (e.key === "Escape") {
        onClose();
      }
    },
    [searchText, result, performSearch, handleFindPrevious, handleFindNext, onClose],
  );

  const handleSearchBlur = useCallback(() => {
    if (searchText.trim() && !result) {
      performSearch();
    }
  }, [performSearch, result, searchText]);

  const handleOverlayClick = useCallback((e: React.MouseEvent) => {
    if (e.target === e.currentTarget) {
      // Don't close on overlay click - this is a non-modal dialog
    }
  }, []);

  const handleDialogKeyDown = useCallback(
    (e: KeyboardEvent<HTMLDivElement>) => {
      if (e.key === "Escape") {
        onClose();
      }
    },
    [onClose],
  );

  if (!isOpen) {
    return null;
  }

  const hasMatches = result && result.totalCount > 0;
  const noMatches = result && result.totalCount === 0 && searchText.trim();
  const overlayStyle = getFindReplaceOverlayStyle(style);
  const titleId = `${id}-find-replace-dialog-title`;
  const findTextId = `${id}-find-text`;

  return (
    <div
      role="presentation"
      className={`docx-find-replace-dialog-overlay pointer-events-none fixed end-0 bottom-0 z-[10002] flex items-start justify-end bg-transparent ${className || ""}`}
      style={overlayStyle}
      data-slot="folio-find-replace-overlay"
      onClick={handleOverlayClick}
      onKeyDown={handleDialogKeyDown}
    >
      <div
        className="docx-find-replace-dialog bg-popover text-popover-foreground pointer-events-auto me-4 mt-3 w-[min(440px,calc(100vw-var(--folio-find-replace-left,5.5rem)-2rem))] rounded-lg border shadow-xl"
        data-testid="find-replace-dialog"
        role="dialog"
        aria-modal="false"
        aria-labelledby={titleId}
      >
        <div className="bg-muted/30 flex items-center justify-between gap-3 border-b px-3 py-2">
          <h2 className="flex min-w-0 items-center gap-2 text-sm font-medium" id={titleId}>
            <SearchIcon className="text-muted-foreground size-4 shrink-0" />
            <span className="truncate">{t("findReplace.find")}</span>
          </h2>
          <Button
            onClick={onClose}
            aria-label={t("findReplace.close")}
            size="icon-xs"
            title={t("findReplace.close")}
            variant="ghost"
          >
            <XIcon />
          </Button>
        </div>

        <div className="space-y-2 p-3">
          <div className="grid grid-cols-[4.5rem_minmax(0,1fr)_auto] items-center gap-2">
            <label className="text-muted-foreground text-xs font-medium" htmlFor={findTextId}>
              {t("findReplace.find")}
            </label>
            <Input
              ref={searchInputRef}
              id={findTextId}
              nativeInput
              type="text"
              className="h-8"
              size="sm"
              value={searchText}
              onChange={handleSearchChange}
              onKeyDown={handleSearchKeyDown}
              onBlur={handleSearchBlur}
              placeholder={t("findReplace.findPlaceholder")}
              aria-label={t("findReplace.findText")}
            />
            <div className="flex items-center gap-0.5">
              <Button
                onClick={handleFindPrevious}
                disabled={!hasMatches}
                aria-label={t("findReplace.previous")}
                title={t("findReplace.previousShortcut")}
                size="icon-xs"
                variant="ghost"
              >
                <ChevronUpIcon />
              </Button>
              <Button
                onClick={handleFindNext}
                disabled={!hasMatches}
                aria-label={t("findReplace.next")}
                title={t("findReplace.nextShortcut")}
                size="icon-xs"
                variant="ghost"
              >
                <ChevronDownIcon />
              </Button>
            </div>
          </div>

          {hasMatches && (
            <div className="text-muted-foreground ms-20 text-xs tabular-nums">
              {t("findReplace.matchCounter", {
                current: String(result.currentIndex + 1),
                total: String(result.totalCount),
              })}
            </div>
          )}
          {noMatches && (
            <div className="text-destructive ms-20 text-xs">{t("findReplace.noResults")}</div>
          )}

          <div className="ms-20 flex flex-wrap items-center gap-x-4 gap-y-2">
            <label className="text-muted-foreground flex items-center gap-2 text-xs">
              <Checkbox checked={matchCase} onCheckedChange={handleMatchCaseChange} />
              {t("findReplace.matchCase")}
            </label>
            <label className="text-muted-foreground flex items-center gap-2 text-xs">
              <Checkbox checked={matchWholeWord} onCheckedChange={handleMatchWholeWordChange} />
              {t("findReplace.wholeWords")}
            </label>
          </div>
        </div>
      </div>
    </div>
  );
}
