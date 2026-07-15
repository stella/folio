import { computed, reactive, ref, shallowRef } from "vue";
import type { ComputedRef, Ref, ShallowRef } from "vue";
import { TextSelection } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";

import { FindReplaceManager } from "@stll/folio-core/managers/FindReplaceManager";
import {
  findInProseMirrorDocument,
  type ProseMirrorFindMatch,
} from "@stll/folio-core/prosemirror/findReplaceSelection";
import { createDefaultFindOptions, type FindOptions } from "@stll/folio-core/utils/findReplace";

export type UseFindReplaceOptions = {
  editorView: Readonly<Ref<EditorView | null>>;
  scrollVisiblePositionIntoView?: (pmPos: number) => void;
};

export type UseFindReplaceReturn = {
  searchText: Ref<string>;
  replaceText: Ref<string>;
  options: FindOptions;
  matches: ShallowRef<ProseMirrorFindMatch[]>;
  currentIndex: Ref<number>;
  currentMatch: ComputedRef<ProseMirrorFindMatch | null>;
  performSearch: () => ProseMirrorFindMatch[];
  goToMatch: (index: number) => boolean;
  findNext: () => ProseMirrorFindMatch | null;
  findPrevious: () => ProseMirrorFindMatch | null;
  replaceCurrent: () => boolean;
  replaceAll: () => number;
  clear: () => void;
};

export function useFindReplace({
  editorView,
  scrollVisiblePositionIntoView,
}: UseFindReplaceOptions): UseFindReplaceReturn {
  const manager = new FindReplaceManager<ProseMirrorFindMatch>();
  const searchText = ref("");
  const replaceText = ref("");
  const options = reactive(createDefaultFindOptions());
  const matches = shallowRef<ProseMirrorFindMatch[]>([]);
  const currentIndex = ref(-1);
  const currentMatch = computed(() => matches.value.at(currentIndex.value) ?? null);

  const clear = (): void => {
    manager.clear();
    matches.value = [];
    currentIndex.value = -1;
  };

  const revealMatch = (match: ProseMirrorFindMatch, index: number): boolean => {
    const view = editorView.value;
    if (!view || match.to > view.state.doc.content.size) {
      return false;
    }
    currentIndex.value = index;
    view.dispatch(view.state.tr.setSelection(createTextSelection(view, match)));
    scrollVisiblePositionIntoView?.(match.from);
    return true;
  };

  const performSearch = (): ProseMirrorFindMatch[] => {
    const view = editorView.value;
    if (!view || !searchText.value.trim()) {
      clear();
      return [];
    }
    const nextMatches = findInProseMirrorDocument(view.state.doc, searchText.value, options);
    manager.setMatches(nextMatches);
    matches.value = nextMatches;
    currentIndex.value = nextMatches.length > 0 ? 0 : -1;
    const firstMatch = nextMatches.at(0);
    if (firstMatch) {
      revealMatch(firstMatch, 0);
    }
    return nextMatches;
  };

  const goToMatch = (index: number): boolean => {
    const selected = manager.goTo(index);
    return selected ? revealMatch(selected.match, selected.index) : false;
  };

  const navigate = (direction: "next" | "previous"): ProseMirrorFindMatch | null => {
    const stepped = manager.navigate(direction);
    if (!stepped || !revealMatch(stepped.match, stepped.index)) {
      return null;
    }
    return stepped.match;
  };

  const replaceCurrent = (): boolean => {
    const view = editorView.value;
    const match = currentMatch.value;
    if (!view || !match || match.to > view.state.doc.content.size) {
      return false;
    }
    view.dispatch(view.state.tr.insertText(replaceText.value, match.from, match.to));
    performSearch();
    return true;
  };

  const replaceAll = (): number => {
    const view = editorView.value;
    if (!view || matches.value.length === 0) {
      return 0;
    }
    let transaction = view.state.tr;
    const sortedMatches = [...matches.value].toSorted((a, b) => b.from - a.from);
    for (const match of sortedMatches) {
      transaction = transaction.insertText(replaceText.value, match.from, match.to);
    }
    view.dispatch(transaction);
    const replacedCount = sortedMatches.length;
    performSearch();
    return replacedCount;
  };

  return {
    searchText,
    replaceText,
    options,
    matches,
    currentIndex,
    currentMatch,
    performSearch,
    goToMatch,
    findNext: () => navigate("next"),
    findPrevious: () => navigate("previous"),
    replaceCurrent,
    replaceAll,
    clear,
  };
}

const createTextSelection = (view: EditorView, match: ProseMirrorFindMatch) => {
  const $from = view.state.doc.resolve(match.from);
  const $to = view.state.doc.resolve(match.to);
  return TextSelection.between($from, $to);
};
