import { computed, onBeforeUnmount, reactive, ref, shallowRef } from "vue";
import type { Ref } from "vue";
import type { EditorView } from "prosemirror-view";

import { FindReplaceManager } from "@stll/folio-core/managers/FindReplaceManager";
import type { DisplayMode } from "@stll/folio-core/managers/EditorModeManager";
import { resolveFindMatchRange } from "@stll/folio-core/prosemirror/findReplaceSelection";
import type { Document } from "@stll/folio-core/types/document";
import type { Layout } from "@stll/folio-core/layout-engine";
import { createFindSearchKey, sameFindSearchKey } from "@stll/folio-core/utils/findSearchKey";
import {
  createFindMatchView,
  getRenderedSearchStories,
} from "@stll/folio-core/utils/findMatchView";
import {
  createDefaultFindOptions,
  findInDocument,
  replaceEligibility,
  type FindMatch,
  type FindOptions,
  type FindResult,
  type ReplaceAllSummary,
} from "@stll/folio-core/utils/findReplace";

export type UseFindReplaceOptions = {
  editorView: Readonly<Ref<EditorView | null>>;
  pagesRef: Readonly<Ref<HTMLElement | null>>;
  getDocument: () => Document | null;
  layout: Readonly<Ref<Layout | null>>;
  commentRoot?: Readonly<Ref<HTMLElement | null>>;
  revealComment?: (commentId: number) => void;
  displayMode: Readonly<Ref<DisplayMode>>;
  readOnly: Readonly<Ref<boolean>>;
  onResultChange: (result: FindResult | null) => void;
};

export function useFindReplace({
  editorView,
  pagesRef,
  getDocument,
  layout,
  commentRoot,
  revealComment,
  displayMode,
  readOnly,
  onResultChange,
}: UseFindReplaceOptions) {
  const manager = new FindReplaceManager<FindMatch>();
  const searchText = ref("");
  const replaceText = ref("");
  const options = reactive(createDefaultFindOptions());
  const matches = shallowRef<FindMatch[]>([]);
  const currentIndex = ref(-1);
  const currentMatch = computed(() => matches.value.at(currentIndex.value) ?? null);
  const canReplaceCurrent = computed(() => {
    const match = currentMatch.value;
    return (
      !readOnly.value &&
      !!editorView.value &&
      !!match &&
      isSearchCurrent() &&
      replaceEligibility(match).status === "allowed"
    );
  });
  let matchView: ReturnType<typeof createFindMatchView> | null = null;
  let searchTimer: ReturnType<typeof setTimeout> | undefined;
  let completedSearchKey: ReturnType<typeof createFindSearchKey> | null = null;

  const searchInputFor = (document: Document) => {
    const bodyDoc = editorView.value?.state.doc;
    return {
      query: searchText.value,
      options: { ...options, displayMode: displayMode.value, includeComments: !!commentRoot },
      document,
      ...(bodyDoc ? { bodyVersion: bodyDoc } : {}),
      context: { renderedStories: getRenderedSearchStories(layout.value) },
    };
  };
  const isSearchCurrent = () => {
    const document = getDocument();
    return (
      !!document &&
      sameFindSearchKey(completedSearchKey, createFindSearchKey(searchInputFor(document)))
    );
  };

  const getFindResult = (): FindResult | null => manager.getResult();
  const notify = () => onResultChange(getFindResult());
  const cancelSearch = () => {
    clearTimeout(searchTimer);
    searchTimer = undefined;
  };
  const clear = (): void => {
    cancelSearch();
    manager.clear();
    matchView?.destroy();
    matchView = null;
    matches.value = [];
    completedSearchKey = null;
    currentIndex.value = -1;
    searchText.value = "";
    notify();
  };

  const revealMatch = (index: number): boolean => {
    currentIndex.value = index;
    matchView?.navigate(index, revealComment);
    notify();
    return true;
  };

  const performSearch = (cause: "query" | "refresh" = "query"): FindMatch[] => {
    cancelSearch();
    const document = getDocument();
    const container = pagesRef.value;
    const bodyDoc = editorView.value?.state.doc;
    if (!document || !searchText.value.trim()) {
      clear();
      return [];
    }
    const input = searchInputFor(document);
    const nextSearchKey = createFindSearchKey(input);
    const refresh = cause === "refresh" && completedSearchKey?.query === nextSearchKey.query;
    const nextMatches = findInDocument(input.document, input.query, input.options, input.context);
    const result = manager.setMatches(
      nextMatches,
      refresh ? (manager.getResult()?.currentIndex ?? 0) : 0,
    );
    completedSearchKey = nextSearchKey;
    matches.value = nextMatches;
    currentIndex.value = nextMatches.length > 0 ? result.currentIndex : -1;
    matchView?.destroy();
    matchView = container
      ? createFindMatchView({
          container,
          document,
          bodyDoc,
          matches: nextMatches,
          commentRoot: commentRoot?.value ?? undefined,
          layout: layout.value,
        })
      : null;
    matchView?.paint(currentIndex.value);
    const firstMatch = nextMatches.at(0);
    if (firstMatch && !refresh) revealMatch(0);
    else notify();
    return nextMatches;
  };

  const scheduleSearch = () => {
    cancelSearch();
    searchTimer = setTimeout(() => performSearch("refresh"), 120);
  };
  const find = (text: string, nextOptions?: Partial<FindOptions>): FindResult | null => {
    searchText.value = text;
    Object.assign(options, createDefaultFindOptions(), nextOptions);
    performSearch();
    return getFindResult();
  };
  const goToMatch = (index: number): boolean => {
    const selected = manager.goTo(index);
    return selected ? revealMatch(selected.index) : false;
  };
  const navigate = (direction: "next" | "previous"): FindMatch | null => {
    const stepped = manager.navigate(direction);
    if (!stepped) return null;
    revealMatch(stepped.index);
    return stepped.match;
  };

  const refreshStaleResults = (): boolean => {
    if (isSearchCurrent()) return false;
    performSearch("refresh");
    return true;
  };

  // Replacement remains an editing command on the main editor. Story results
  // must never be interpreted as positions in the body document.
  const replaceCurrent = (): boolean => {
    const view = editorView.value;
    if (!view || readOnly.value || refreshStaleResults()) return false;
    const match = currentMatch.value;
    if (!canReplaceCurrent.value || !view || !match) return false;
    if (match.text === replaceText.value) return true;
    const range = resolveFindMatchRange(view.state.doc, match);
    if (!range) return false;
    view.dispatch(view.state.tr.insertText(replaceText.value, range.from, range.to));
    performSearch();
    return true;
  };
  const replaceAll = (): ReplaceAllSummary => {
    const view = editorView.value;
    const totalCount = matches.value.length;
    if (readOnly.value || !view || refreshStaleResults())
      return { replacedCount: 0, skippedCount: totalCount };
    const eligibleMatches = matches.value.filter(
      (match) => replaceEligibility(match).status === "allowed",
    );
    const unchangedCount = eligibleMatches.filter(
      (match) => match.text === replaceText.value,
    ).length;
    const ranges = eligibleMatches
      .filter((match) => match.text !== replaceText.value)
      .flatMap((match) => {
        const range = resolveFindMatchRange(view.state.doc, match);
        return range ? [range] : [];
      })
      .toSorted((a, b) => b.from - a.from);
    const replacedCount = ranges.length + unchangedCount;
    const summary = { replacedCount, skippedCount: totalCount - replacedCount };
    if (!ranges.length) return summary;
    let transaction = view.state.tr;
    for (const range of ranges)
      transaction = transaction.insertText(replaceText.value, range.from, range.to);
    view.dispatch(transaction);
    performSearch();
    return summary;
  };

  onBeforeUnmount(() => {
    cancelSearch();
    matchView?.destroy();
  });
  return {
    searchText,
    replaceText,
    options,
    matches,
    currentIndex,
    currentMatch,
    canReplaceCurrent,
    performSearch,
    scheduleSearch,
    find,
    getFindResult,
    goToMatch,
    findNext: () => navigate("next"),
    findPrevious: () => navigate("previous"),
    replaceCurrent,
    replaceAll,
    clear,
  };
}

export type UseFindReplaceReturn = ReturnType<typeof useFindReplace>;
