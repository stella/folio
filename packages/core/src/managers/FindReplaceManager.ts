/**
 * FindReplaceManager
 *
 * Framework-agnostic find/replace result cursor and replace orchestration.
 * Extracted from the React `useFindReplace` hook (the document-aware half).
 *
 * Owns the active find result (the matches plus the cursor index) and runs the
 * document-level replace operations. Matching itself (scanning a document for
 * hits) and the side effects (scroll-into-view, host selection, dialog state)
 * stay in the adapter. The manager is generic over the adapter's match shape
 * and only reads each match's paragraph-relative offsets via `FindMatchPosition`,
 * so a richer adapter match (content index, matched text) round-trips untouched.
 */

import type { FindMatchPosition } from "../prosemirror/findReplaceSelection";
import type { Document } from "../types/document";
import { replaceTextInDocument } from "../utils/replaceText";

export type FindDirection = "next" | "previous";

export type FindReplaceResult<TMatch> = {
  matches: TMatch[];
  totalCount: number;
  currentIndex: number;
};

export type ReplaceAllOutcome = {
  document: Document;
  replacedCount: number;
};

/**
 * The cursor index after stepping `direction` over `totalCount` matches,
 * wrapping at both ends. Returns 0 when there are no matches.
 */
export const getAdjacentFindIndex = (
  currentIndex: number,
  totalCount: number,
  direction: FindDirection,
): number => {
  if (totalCount <= 0) {
    return 0;
  }

  if (direction === "previous") {
    return currentIndex === 0 ? totalCount - 1 : currentIndex - 1;
  }

  return (currentIndex + 1) % totalCount;
};

const toReplaceRange = (match: FindMatchPosition) => ({
  start: { paragraphIndex: match.paragraphIndex, offset: match.startOffset },
  end: { paragraphIndex: match.paragraphIndex, offset: match.endOffset },
});

export class FindReplaceManager<TMatch extends FindMatchPosition> {
  private result: FindReplaceResult<TMatch> | null = null;

  /** The active find result, or null when there is no active search. */
  getResult(): FindReplaceResult<TMatch> | null {
    return this.result;
  }

  /** Discard the active search. */
  clear(): void {
    this.result = null;
  }

  /**
   * Store a fresh set of matches as the active result, with the cursor on the
   * first match. The result is kept even when `matches` is empty so callers can
   * surface a "no results" state.
   */
  setMatches(matches: TMatch[]): FindReplaceResult<TMatch> {
    this.result = { matches, totalCount: matches.length, currentIndex: 0 };
    return this.result;
  }

  /**
   * Step the cursor and return the now-current match with its index, or null
   * when there are no matches to navigate.
   */
  navigate(direction: FindDirection): { match: TMatch; index: number } | null {
    const result = this.result;
    if (!result || result.matches.length === 0) {
      return null;
    }

    const index = getAdjacentFindIndex(result.currentIndex, result.matches.length, direction);
    this.result = { ...result, currentIndex: index };
    const match = result.matches[index];
    return match ? { match, index } : null;
  }

  /**
   * Move the cursor to an explicit match index. Returns the selected match, or
   * null when there is no active result or the index is out of bounds.
   */
  goTo(index: number): { match: TMatch; index: number } | null {
    const result = this.result;
    if (!result || index < 0 || index >= result.matches.length) {
      return null;
    }
    const match = result.matches.at(index);
    if (!match) {
      return null;
    }
    this.result = { ...result, currentIndex: index };
    return { match, index };
  }

  /**
   * Replace the current match's text. Returns the new document, or null when
   * there is no current match or the replace fails.
   */
  replaceCurrent(document: Document, replaceText: string): Document | null {
    const result = this.result;
    if (!result || result.matches.length === 0) {
      return null;
    }

    const match = result.matches[result.currentIndex];
    if (!match) {
      return null;
    }

    try {
      return replaceTextInDocument(document, toReplaceRange(match), replaceText);
    } catch {
      return null;
    }
  }

  /**
   * Replace every match. Matches are replaced from the end of the document
   * backwards so earlier offsets stay valid as the text shifts. Clears the
   * active result and returns the new document plus the match count, or null
   * when there is nothing to replace.
   */
  replaceAll(
    document: Document,
    matches: readonly TMatch[],
    replaceText: string,
  ): ReplaceAllOutcome | null {
    if (matches.length === 0) {
      return null;
    }

    const sortedMatches = [...matches].toSorted((a, b) => {
      if (a.paragraphIndex !== b.paragraphIndex) {
        return b.paragraphIndex - a.paragraphIndex;
      }
      return b.startOffset - a.startOffset;
    });

    let nextDocument = document;
    for (const match of sortedMatches) {
      try {
        nextDocument = replaceTextInDocument(nextDocument, toReplaceRange(match), replaceText);
      } catch {
        continue;
      }
    }

    this.result = null;
    return { document: nextDocument, replacedCount: matches.length };
  }
}
