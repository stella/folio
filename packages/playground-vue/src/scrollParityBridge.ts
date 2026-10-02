import type {
  DocxEditorRef,
  DocxEditorProps,
  FolioSuggestion,
  FolioAIEditOperation,
  FolioAIEditSnapshot,
} from "@stll/folio-vue";

type FolioDocumentOperationBatch = Parameters<DocxEditorRef["applyDocumentOperations"]>[0]["batch"];
type EditorView = Parameters<NonNullable<DocxEditorProps["onEditorViewReady"]>>[0];
type PagedRef = NonNullable<ReturnType<DocxEditorRef["getEditorRef"]>>;
type ScrollMethod =
  | Extract<keyof DocxEditorRef, `scrollTo${string}`>
  | "highlightPassage"
  | "showInDocument";
type PagedScrollMethod = Extract<keyof PagedRef, `scrollTo${string}`>;

// Total over both public ref contracts: adding a navigation API requires a
// browser case here. None of the public scrolling methods is exempt.
export const SCROLL_NAVIGATION_CASES = {
  scrollToPage: { method: "scrollToPage" },
  scrollToParaId: { method: "scrollToParaId" },
  scrollToAIEditOperation: { method: "scrollToAIEditOperation" },
  scrollToBlock: { method: "scrollToBlock" },
  scrollToSuggestion: { method: "scrollToSuggestion" },
  scrollToContentControl: { method: "scrollToContentControl" },
  highlightPassage: { method: "highlightPassage" },
  showInDocument: { method: "showInDocument" },
} as const satisfies { [Method in ScrollMethod]: { method: Method } };
export const PAGED_SCROLL_NAVIGATION_CASES = {
  scrollToPage: { method: "scrollToPage" },
  scrollToParaId: { method: "scrollToParaId" },
  scrollToPosition: { method: "scrollToPosition" },
} as const satisfies { [Method in PagedScrollMethod]: { method: Method } };

export const SCROLL_TARGET_PARA_ID = "13300300";
export const SCROLL_TARGET_TEXT = "Scroll destination on page three";
export const SCROLL_REVISION_ID = 1330;
export const SCROLL_CONTROL_TAG = "scroll-root-target";

export const SCROLL_TARGET_SUGGESTION = {
  id: "scroll-target-suggestion",
  type: "insertAfterBlock",
  blockId: SCROLL_TARGET_PARA_ID,
  text: "Suggested destination",
} as const satisfies FolioAIEditOperation;

type FindScrollSuggestionTargetOptions = {
  root: ParentNode;
  snapshot: FolioAIEditSnapshot;
  suggestion: FolioSuggestion;
};

// Insertions may inherit pageBreakBefore and appear on a different page from
// their source block. Resolve the live suggestion range before measuring it.
export const findScrollSuggestionTarget = ({
  root,
  snapshot,
  suggestion,
}: FindScrollSuggestionTargetOptions) => {
  const range = suggestion.ranges.at(0);
  if (!range) return null;
  const block = snapshot.blocks.find(({ id }) => {
    const anchor = snapshot.anchors[id];
    return anchor !== undefined && anchor.from >= range.from && anchor.from < range.to;
  });
  const anchor = block ? snapshot.anchors[block.id] : undefined;
  return anchor
    ? root.querySelector<HTMLElement>(`.layout-paragraph[data-pm-start="${anchor.from}"]`)
    : null;
};

export type ScrollParityBridge = ReturnType<typeof buildScrollParityBridge>;

export const buildScrollParityBridge = (getRef: () => DocxEditorRef | null) => {
  let suggestionId: string | null = null;
  let readyCount = 0;
  let readyAppliedTop = 0;
  let readyEventCount = 0;
  let readyEventAppliedTop = 0;
  const requireRef = () => {
    const ref = getRef();
    if (!ref) throw new Error("Host flow requires a document ref");
    return ref;
  };
  return {
    loadFlowDocument: async (source: number[]) => {
      const ref = requireRef();
      await ref.loadDocumentBuffer(new Uint8Array(source));
      ref.ensureEditorView({ focus: false });
      suggestionId = null;
    },
    replaceFlowDocument: async (source: number[]) => {
      const ref = requireRef();
      // Parse a distinct replacement, then exercise the pre-parsed public API.
      await ref.loadDocumentBuffer(new Uint8Array(source));
      const document = ref.getDocument();
      if (!document) throw new Error("Replacement document unavailable");
      ref.loadDocument(document);
      ref.ensureEditorView({ focus: false });
      suggestionId = null;
    },
    applyFlowBatch: (batch: FolioDocumentOperationBatch) => {
      const ref = requireRef();
      const snapshot = ref.createAIEditSnapshot();
      if (!snapshot) throw new Error("Host flow snapshot unavailable");
      return ref.applyDocumentOperations({ snapshot, batch });
    },
    readFlowDocument: () => {
      const ref = requireRef();
      const document = ref.getDocument();
      const pagedDocument = ref.getEditorRef()?.getDocument();
      if (!document || !pagedDocument) throw new Error("Host document read unavailable");
      return {
        documentBody: document.package.document,
        pagedBody: pagedDocument.package.document,
        liveText: ref.getEditor()?.getState()?.doc.textContent,
      };
    },
    saveFlowDocument: async () => {
      const saved = await requireRef().save();
      if (!saved) throw new Error("Host checkpoint did not save");
      return [...new Uint8Array(saved)];
    },
    resetFlowScroll: () => {
      const root = requireRef().getScrollRoot();
      if (!root) throw new Error("Host scroll root unavailable");
      root.scrollTop = 0;
    },
    rejectFlowSuggestion: () => {
      const ref = requireRef();
      if (!suggestionId) return false;
      const rejected = ref.rejectSuggestion(suggestionId);
      suggestionId = null;
      return rejected;
    },
    prepareSuggestion: () => {
      const ref = getRef();
      const snapshot = ref?.createAIEditSnapshot();
      if (!ref || !snapshot) return false;
      const result = ref.applyAIEditOperations({
        snapshot,
        mode: "suggested",
        operations: [SCROLL_TARGET_SUGGESTION],
      });
      suggestionId = result.applied.at(0)?.suggestionId ?? null;
      return suggestionId !== null;
    },
    navigate: (method: ScrollMethod) => {
      const ref = getRef();
      if (!ref) return false;
      switch (method) {
        case "scrollToPage":
          ref.scrollToPage(3);
          return true;
        case "scrollToParaId":
          return ref.scrollToParaId(SCROLL_TARGET_PARA_ID);
        case "scrollToAIEditOperation":
          return ref.scrollToAIEditOperation(SCROLL_REVISION_ID);
        case "scrollToBlock":
          return ref.scrollToBlock(SCROLL_TARGET_PARA_ID);
        case "scrollToSuggestion":
          return suggestionId !== null && ref.scrollToSuggestion(suggestionId);
        case "scrollToContentControl":
          return ref.scrollToContentControl({ tag: SCROLL_CONTROL_TAG });
        case "highlightPassage":
          return (
            ref.highlightPassage({ blockId: SCROLL_TARGET_PARA_ID, text: SCROLL_TARGET_TEXT }) ===
            "passage"
          );
        case "showInDocument":
          return ref.showInDocument({
            type: "block",
            story: "main",
            blockId: SCROLL_TARGET_PARA_ID,
          });
        default: {
          const exhaustive: never = method;
          return exhaustive;
        }
      }
    },
    navigatePaged: (method: PagedScrollMethod) => {
      const ref = getRef();
      const paged = ref?.getEditorRef();
      if (!ref || !paged) return false;
      switch (method) {
        case "scrollToPage":
          paged.scrollToPage(3);
          return true;
        case "scrollToParaId":
          return paged.scrollToParaId(SCROLL_TARGET_PARA_ID);
        case "scrollToPosition": {
          const anchor = ref.createAIEditSnapshot()?.anchors[SCROLL_TARGET_PARA_ID];
          if (!anchor) return false;
          paged.scrollToPosition(anchor.from);
          return true;
        }
        default: {
          const exhaustive: never = method;
          return exhaustive;
        }
      }
    },
    readTarget: (method?: ScrollMethod) => {
      const ref = getRef();
      const root = ref?.getScrollRoot();
      if (!ref || !root) return null;
      const paraId = method === "scrollToContentControl" ? "13300301" : SCROLL_TARGET_PARA_ID;
      const resolveTarget = () => {
        if (method !== "scrollToSuggestion") {
          return root.querySelector<HTMLElement>(`.layout-paragraph[data-para-id="${paraId}"]`);
        }
        const suggestion = ref
          .getSuggestions()
          .find((candidate) => candidate.suggestionId === suggestionId);
        const snapshot = ref.createAIEditSnapshot();
        if (!suggestion || !snapshot) return null;
        return findScrollSuggestionTarget({ root, snapshot, suggestion });
      };
      const target = resolveTarget();
      if (!target) return null;
      const targetRect = target.getBoundingClientRect();
      const rootRect = root.getBoundingClientRect();
      return {
        scrollTop: root.scrollTop,
        top: targetRect.top,
        bottom: targetRect.bottom,
        viewportTop: rootRect.top,
        viewportBottom: rootRect.bottom,
      };
    },
    onViewReady: (view: EditorView) => {
      if (!view) return;
      const requestedTop = Number(new URLSearchParams(location.search).get("readyScroll"));
      if (!(requestedTop > 0)) return;
      const root =
        view.dom.closest<HTMLElement>("[data-folio-scroll]") ?? getRef()?.getScrollRoot();
      if (!root) return;
      readyCount++;
      root.scrollTop = requestedTop;
      readyAppliedTop = root.scrollTop;
    },
    onReady: () => {
      if (!new URLSearchParams(location.search).has("readyEventScroll")) return;
      const root = getRef()?.getScrollRoot();
      if (!root) return;
      readyEventCount++;
      root.scrollTop = 500;
      readyEventAppliedTop = root.scrollTop;
    },
    readReady: () => ({
      count: readyCount,
      appliedTop: readyAppliedTop,
      eventCount: readyEventCount,
      eventAppliedTop: readyEventAppliedTop,
      scrollTop: getRef()?.getScrollRoot()?.scrollTop ?? null,
    }),
    reloadForReady: async () => {
      const ref = getRef();
      const doc = ref?.getDocument();
      if (!ref || !doc?.originalBuffer) return false;
      await ref.loadDocumentBuffer(doc.originalBuffer);
      return true;
    },
  };
};
