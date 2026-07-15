import { prefersReducedMotionBehavior } from "@stll/folio-core/paged-layout/scrollNavigation";
import {
  createDefaultFindOptions,
  createSearchPattern,
  escapeRegexString,
  findAllMatches,
  findInDocument,
  findInParagraph,
} from "@stll/folio-core/utils/findReplace";
import type { FindMatch, FindOptions, FindResult } from "@stll/folio-core/utils/findReplace";

export type { FindMatch, FindOptions, FindResult };
export {
  createDefaultFindOptions,
  createSearchPattern,
  escapeRegexString,
  findAllMatches,
  findInDocument,
  findInParagraph,
};

export type HighlightOptions = {
  currentMatchColor: string;
  otherMatchColor: string;
};

export const replaceAllInContent = (
  content: string,
  searchText: string,
  replaceText: string,
  options: FindOptions,
): string => {
  const pattern = createSearchPattern(searchText, options);
  return pattern ? content.replace(pattern, replaceText) : content;
};

export const replaceFirstInContent = (
  content: string,
  searchText: string,
  replaceText: string,
  options: FindOptions,
  startIndex: number = 0,
): { content: string; replaced: boolean; matchStart: number; matchEnd: number } => {
  const matches = findAllMatches(content, searchText, options);
  const match = matches.find(({ start }) => start >= startIndex) ?? matches.at(0);
  if (!match) {
    return { content, replaced: false, matchStart: -1, matchEnd: -1 };
  }
  return {
    content: content.slice(0, match.start) + replaceText + content.slice(match.end),
    replaced: true,
    matchStart: match.start,
    matchEnd: match.start + replaceText.length,
  };
};

export const getMatchCountText = (result: FindResult | null): string => {
  if (!result) {
    return "";
  }
  if (result.totalCount === 0) {
    return "No results";
  }
  if (result.totalCount === 1) {
    return "1 match";
  }
  return `${result.currentIndex + 1} of ${result.totalCount} matches`;
};

export const isEmptySearch = (searchText: string): boolean =>
  !searchText || searchText.trim() === "";

export const getDefaultHighlightOptions = (): HighlightOptions => ({
  currentMatchColor: "var(--doc-find-current)",
  otherMatchColor: "var(--doc-find-highlight)",
});

export const scrollToMatch = (containerElement: HTMLElement | null, match: FindMatch): void => {
  if (!containerElement) {
    return;
  }
  const paragraphElement =
    containerElement.querySelector(`[data-paragraph-index="${match.paragraphIndex}"]`) ??
    containerElement.querySelector(
      `.layout-paragraph[data-block-id="block-${match.paragraphIndex + 1}"]`,
    ) ??
    containerElement.querySelectorAll(".layout-paragraph").item(match.paragraphIndex);
  paragraphElement.scrollIntoView({
    behavior: prefersReducedMotionBehavior(),
    block: "center",
  });
};
