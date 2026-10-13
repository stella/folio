import { computed, nextTick, onBeforeUnmount, reactive, ref, shallowRef } from "vue";
import type { Ref } from "vue";
import type { EditorView } from "prosemirror-view";

import { FindReplaceManager } from "@stll/folio-core/managers/FindReplaceManager";
import type { DisplayMode } from "@stll/folio-core/managers/EditorModeManager";
import { resolveFindMatchRange } from "@stll/folio-core/prosemirror/findReplaceSelection";
import type { Document } from "@stll/folio-core/types/document";
import type { Layout } from "@stll/folio-core/layout-engine";
import { fromProseDoc } from "@stll/folio-core/prosemirror/conversion/fromProseDoc";
import {
  createFindMatchView,
  getRenderedSearchStories,
} from "@stll/folio-core/utils/findMatchView";
import {
  createDefaultFindOptions,
  findInDocument,
  type FindMatch,
  type FindOptions,
  type FindResult,
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
  let matchView: ReturnType<typeof createFindMatchView> | null = null;
  let searchTimer: ReturnType<typeof setTimeout> | undefined;
  let matchedQuery = "";

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
    currentIndex.value = -1;
    searchText.value = "";
    notify();
  };

  const revealMatch = (match: FindMatch, index: number): boolean => {
    currentIndex.value = index;
    matchView?.paint(index);
    matchView?.reveal(index);
    if (match.commentId !== undefined) {
      revealComment?.(match.commentId);
      void nextTick(() => {
        if (matches.value.at(index) !== match || currentIndex.value !== index) return;
        matchView?.paint(index);
        matchView?.reveal(index);
      });
    }
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
    const refresh = cause === "refresh" && matchedQuery === searchText.value;
    matchedQuery = searchText.value;
    const searchDocument = bodyDoc
      ? fromProseDoc(bodyDoc, document, { stylesheetSource: { type: "package" } })
      : document;
    const nextMatches = findInDocument(
      searchDocument,
      searchText.value,
      {
        ...options,
        displayMode: displayMode.value,
        includeComments: !!commentRoot,
      },
      { renderedStories: getRenderedSearchStories(layout.value) },
    );
    const result = manager.setMatches(
      nextMatches,
      refresh ? (manager.getResult()?.currentIndex ?? 0) : 0,
    );
    matches.value = nextMatches;
    currentIndex.value = nextMatches.length > 0 ? result.currentIndex : -1;
    matchView?.destroy();
    matchView = container
      ? createFindMatchView({
          container,
          document: searchDocument,
          bodyDoc,
          matches: nextMatches,
          commentRoot: commentRoot?.value ?? undefined,
        })
      : null;
    matchView?.paint(currentIndex.value);
    const firstMatch = nextMatches.at(0);
    if (firstMatch && !refresh) revealMatch(firstMatch, 0);
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
    return selected ? revealMatch(selected.match, selected.index) : false;
  };
  const navigate = (direction: "next" | "previous"): FindMatch | null => {
    const stepped = manager.navigate(direction);
    if (!stepped) return null;
    revealMatch(stepped.match, stepped.index);
    return stepped.match;
  };

  // Replacement remains an editing command on the main editor. Story results
  // must never be interpreted as positions in the body document.
  const replaceCurrent = (): boolean => {
    const view = editorView.value;
    const match = currentMatch.value;
    if (
      readOnly.value ||
      !view ||
      !match ||
      match.story !== "main" ||
      match.textBoxIndex !== undefined ||
      match.commentId !== undefined
    )
      return false;
    const range = resolveFindMatchRange(view.state.doc, match);
    if (!range) return false;
    view.dispatch(view.state.tr.insertText(replaceText.value, range.from, range.to));
    performSearch();
    return true;
  };
  const replaceAll = (): number => {
    const view = editorView.value;
    if (
      readOnly.value ||
      !view ||
      matches.value.some(
        (match) =>
          match.story !== "main" ||
          match.textBoxIndex !== undefined ||
          match.commentId !== undefined,
      )
    )
      return 0;
    const ranges = matches.value
      .filter(
        (match) =>
          match.story === "main" &&
          match.textBoxIndex === undefined &&
          match.commentId === undefined,
      )
      .flatMap((match) => {
        const range = resolveFindMatchRange(view.state.doc, match);
        return range ? [range] : [];
      })
      .toSorted((a, b) => b.from - a.from);
    if (!ranges.length) return 0;
    let transaction = view.state.tr;
    for (const range of ranges)
      transaction = transaction.insertText(replaceText.value, range.from, range.to);
    view.dispatch(transaction);
    performSearch();
    return ranges.length;
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
