import type { EditorView } from "prosemirror-view";
import {
  computed,
  onScopeDispose,
  reactive,
  ref,
  shallowRef,
  watch,
} from "vue";
import type { Ref } from "vue";

import type { Layout } from "@stll/folio-core/layout-engine";
import { getEditorReadOnlyReason } from "@stll/folio-core/managers/EditorModeManager";
import type {
  DisplayMode,
  EditorReadOnlyReason,
} from "@stll/folio-core/managers/EditorModeManager";
import { FindReplaceManager } from "@stll/folio-core/managers/FindReplaceManager";
import { findInDisplayedDocument } from "@stll/folio-core/prosemirror/findReplaceSelection";
import {
  replaceFindMatch,
  replaceFindMatches,
} from "@stll/folio-core/prosemirror/replaceFindMatches";
import type { Document } from "@stll/folio-core/types/document";
import {
  createFindMatchView,
  getRenderedSearchStories,
} from "@stll/folio-core/utils/findMatchView";
import {
  createDefaultFindOptions,
  type FindMatch,
  type FindOptions,
  type FindResult,
  type ReplaceAllSummary,
  replaceEligibility,
} from "@stll/folio-core/utils/findReplace";
import {
  createFindSearchKey,
  sameFindSearchKey,
} from "@stll/folio-core/utils/findSearchKey";

export type UseFindReplaceOptions = {
  editorView: Readonly<Ref<EditorView | null>>;
  pagesRef: Readonly<Ref<HTMLElement | null>>;
  getDocument: () => Document | null;
  layout: Readonly<Ref<Layout | null>>;
  commentRoot?: Readonly<Ref<HTMLElement | null>>;
  revealComment?: (commentId: number) => void;
  displayMode: Readonly<Ref<DisplayMode>>;
  readOnly: Readonly<Ref<boolean>>;
  /** Notified when a replacement is refused because editing is not allowed. */
  onReadonlyEditAttempt?: (reason: EditorReadOnlyReason) => void;
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
  onReadonlyEditAttempt,
  onResultChange,
}: UseFindReplaceOptions) {
  const manager = new FindReplaceManager<FindMatch>();
  const searchText = ref("");
  const replaceText = ref("");
  const options = reactive(createDefaultFindOptions());
  const matches = shallowRef<FindMatch[]>([]);
  const currentIndex = ref(-1);
  const currentMatch = computed(
    () => matches.value.at(currentIndex.value) ?? null,
  );
  /** Why replacement is refused: host read-only, viewing mode, or a review view. */
  const getReplacementRefusal = (): EditorReadOnlyReason | null =>
    getEditorReadOnlyReason({
      // The view's editable prop is not reactive; read it on every decision.
      editingMode: editorView.value?.editable === false ? "viewing" : "editing",
      readOnlyProp: readOnly.value,
      displayMode: displayMode.value,
    });
  const replacementRefusal = computed(getReplacementRefusal);
  const canReplaceCurrent = computed(() => {
    const match = currentMatch.value;
    return (
      getReplacementRefusal() === null &&
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
      options: {
        ...options,
        displayMode: displayMode.value,
        includeComments: !!commentRoot,
      },
      document,
      ...(bodyDoc ? { bodyVersion: bodyDoc } : {}),
      context: { renderedStories: getRenderedSearchStories(layout.value) },
    };
  };
  const isSearchCurrent = () => {
    const document = getDocument();
    return (
      !!document &&
      sameFindSearchKey(
        completedSearchKey,
        createFindSearchKey(searchInputFor(document)),
      )
    );
  };

  const getFindResult = (): FindResult | null => manager.getResult();
  const notify = () => onResultChange(getFindResult());
  const cancelSearch = () => {
    clearTimeout(searchTimer);
    searchTimer = undefined;
  };
  const clearResults = (): void => {
    cancelSearch();
    manager.clear();
    matchView?.destroy();
    matchView = null;
    matches.value = [];
    completedSearchKey = null;
    currentIndex.value = -1;
  };
  const clear = (): void => {
    clearResults();
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
    const nextMatches = findInDisplayedDocument({
      document: input.document,
      state: editorView.value?.state ?? null,
      searchText: input.query,
      options: input.options,
      context: input.context,
    });
    if (nextMatches === null) {
      // A review view reads its body from the mounted editor; keep the query
      // so the next refresh repeats this search once the editor exists.
      clearResults();
      notify();
      return [];
    }
    const nextSearchKey = createFindSearchKey(input);
    const refresh =
      cause === "refresh" && completedSearchKey?.query === nextSearchKey.query;
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
  const find = (
    text: string,
    nextOptions?: Partial<FindOptions>,
  ): FindResult | null => {
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

  const refuseReadonlyReplacement = (): boolean => {
    const reason = getReplacementRefusal();
    if (reason === null) return false;
    onReadonlyEditAttempt?.(reason);
    return true;
  };

  // Replacement is an editing command on the main editor, run through the
  // same core owner as every adapter. Story results are never body positions.
  const replaceCurrent = (): boolean => {
    const view = editorView.value;
    if (!view || refuseReadonlyReplacement() || refreshStaleResults())
      return false;
    const match = currentMatch.value;
    if (!canReplaceCurrent.value || !match) return false;
    const before = view.state.doc;
    const replaced = replaceFindMatch(view, match, replaceText.value);
    if (view.state.doc !== before) performSearch("refresh");
    return replaced;
  };
  const replaceAll = (): ReplaceAllSummary => {
    const view = editorView.value;
    const totalCount = matches.value.length;
    if (!view || refuseReadonlyReplacement() || refreshStaleResults())
      return { replacedCount: 0, skippedCount: totalCount };
    const before = view.state.doc;
    const summary = replaceFindMatches(view, matches.value, replaceText.value);
    if (view.state.doc !== before) performSearch();
    return summary;
  };

  // A display-mode switch changes what is searchable; refresh without waiting.
  watch(displayMode, () => {
    if (searchText.value.trim()) performSearch();
  });

  onScopeDispose(() => {
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
    replacementRefusal,
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
