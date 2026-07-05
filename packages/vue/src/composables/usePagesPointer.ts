/**
 * Pages-area pointer composable — owns every mousedown / mousemove /
 * click / dblclick / scroll handler on the pages viewport, plus the
 * incidental state those handlers own: multi-click detection, drag
 * selection, the table quick-insert button, the header/footer
 * double-click editor state, and the page-indicator scroll tracker.
 * Reads `selectedImage` / `imageInteracting` from `useImageActions`
 * and the table-resize bridge from `useTableResize`. The
 * selection-overlay (caret + text-rect) primitive `clearOverlay`
 * still lives in the parent — passed in as a callback.
 */

import { onBeforeUnmount, onMounted, ref, shallowRef, type Ref, type ShallowRef } from "vue";
import { TextSelection, NodeSelection, type Command } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import type { HeaderFooter, BlockContent } from "@stll/folio-core/types/content";
import type { Document, SectionProperties } from "@stll/folio-core/types/document";
import type { Layout } from "@stll/folio-core/layout-engine";
import { findImageElement } from "@stll/folio-core/layout-painter/imageLayout";
import {
  detectTableInsertHover,
  TABLE_INSERT_HIDE_DELAY_MS,
} from "@stll/folio-core/layout-bridge/tableInsertHover";
import {
  createCellDragTracker,
  findCellPosFromPmPos,
} from "@stll/folio-core/prosemirror/cellDragSelection";
import {
  scrollVisiblePositionIntoView as scrollVisiblePositionIntoViewImpl,
  resolvePos as resolvePosImpl,
  selectWord as selectWordImpl,
  selectParagraph as selectParagraphImpl,
} from "../utils/domQueries";
import type { ImageSelectionInfo } from "../components/imageSelectionTypes";
import type { HyperlinkPopupData } from "../components/ui/hyperlinkPopupTypes";
import { useDragAutoScroll } from "./useDragAutoScroll";

type CommandFactory = (...args: readonly unknown[]) => Command;
type Commands = Record<string, CommandFactory>;

type TableResizeApi = {
  tryStartResize: (e: MouseEvent, view: EditorView) => boolean;
  isResizing: Ref<boolean>;
};

export interface TableInsertButton {
  type: "row" | "column";
  x: number;
  y: number;
  cellPmPos: number;
}

export interface HfEditState {
  position: "header" | "footer";
  rId: string | null;
  headerFooter: HeaderFooter | null;
  targetRect: { top: number; left: number; width: number; height: number } | null;
}

export interface ScrollPageInfo {
  currentPage: number;
  totalPages: number;
  visible: boolean;
}

export interface UsePagesPointerOptions {
  editorView: Ref<EditorView | null>;
  pagesRef: Ref<HTMLElement | null>;
  pagesViewportRef: Ref<HTMLElement | null>;
  selectedImage: ShallowRef<ImageSelectionInfo | null>;
  imageInteracting: Ref<boolean>;
  hyperlinkPopupData: Ref<HyperlinkPopupData | null>;
  readOnly: Ref<boolean>;
  zoom: Ref<number>;
  layout: Ref<Layout | null>;
  tableResize: TableResizeApi;
  getCommands: () => Commands;
  getDocument: () => Document | null;
  reLayout: () => void;
  emit: (event: string, ...args: unknown[]) => void;
  clearOverlay: () => void;
  /**
   * Re-mount HF EditorViews when `package.headers/footers` content changes —
   * exposed by `useDocxEditor.syncHfPMs`. Called after every save so the
   * persistent PM points at the new HeaderFooter object. Optional so existing
   * consumers can no-op until they wire it through.
   */
  syncHfPMs?: () => void;
  /** Resolve the persistent EditorView for an HF instance (for click routing). */
  getHfPmView?: (hf: HeaderFooter) => EditorView | null;
  /**
   * Replace the loaded Document — used by HF materialisation to publish a
   * fresh Document object instead of mutating in place. Optional; if absent,
   * callers fall back to in-place mutation + `syncHfPMs()`.
   */
  setDocument?: (doc: Document) => void;
}

const MULTI_CLICK_DELAY = 500;

export interface UsePagesPointerReturn {
  tableInsertButton: Ref<TableInsertButton | null>;
  hfEdit: ShallowRef<HfEditState | null>;
  scrollPageInfo: Ref<ScrollPageInfo>;
  resolvePos: (clientX: number, clientY: number) => number | null;
  setPmSelection: (anchor: number, head?: number) => void;
  scrollVisiblePositionIntoView: (pmPos: number) => void;
  navigateToBookmark: (bookmarkName: string) => void;
  handlePagesMouseDown: (event: MouseEvent) => void;
  handlePagesMouseMove: (event: MouseEvent) => void;
  handlePagesClick: (event: MouseEvent) => void;
  handlePagesDoubleClick: (event: MouseEvent) => void;
  handleTableInsertClick: (event: MouseEvent) => void;
  clearTableInsertTimer: () => void;
  handleHfSave: (content: BlockContent[]) => void;
  handleHfRemove: () => void;
}

export function usePagesPointer(opts: UsePagesPointerOptions): UsePagesPointerReturn {
  // ─── Table quick-action "+" button ──────────────────────────────────────
  const tableInsertButton = ref<TableInsertButton | null>(null);
  let tableInsertHideTimer: ReturnType<typeof setTimeout> | null = null;
  function clearTableInsertTimer() {
    if (tableInsertHideTimer !== null) {
      clearTimeout(tableInsertHideTimer);
      tableInsertHideTimer = null;
    }
  }

  // ─── Inline header/footer editor ────────────────────────────────────────
  // shallowRef so the nested `headerFooter` reference stays identity-equal
  // to the instance in `Document.package.headers/footers`.
  const hfEdit = shallowRef<HfEditState | null>(null);

  // ─── Multi-click detection (double = word, triple = paragraph) ──────────
  let lastClickTime = 0;
  let lastClickPos: number | null = null;
  let clickCount = 0;

  // ─── Drag-to-select ─────────────────────────────────────────────────────
  let isDragging = false;
  let dragAnchor: number | null = null;
  const cellDrag = createCellDragTracker();

  const dragAutoScroll = useDragAutoScroll({
    pagesContainer: opts.pagesRef,
    onScrollExtendSelection: (clientX, clientY) => {
      if (!isDragging || dragAnchor === null) return;
      const pos = resolvePos(clientX, clientY);
      if (pos !== null && pos !== dragAnchor) setPmSelection(dragAnchor, pos);
    },
  });

  // ─── Page-indicator overlay ─────────────────────────────────────────────
  const scrollPageInfo = ref<ScrollPageInfo>({ currentPage: 1, totalPages: 1, visible: false });
  let scrollFadeTimer: ReturnType<typeof setTimeout> | null = null;

  function resolvePos(clientX: number, clientY: number): number | null {
    return resolvePosImpl(opts.pagesRef.value, opts.editorView.value, clientX, clientY);
  }

  /**
   * The PM EditorView every pointer gesture flows through. When HF edit
   * mode is active and the matching persistent HF view exists, that's the
   * "active" view; otherwise (or as a fallback) it's the body PM.
   */
  function activeView(): EditorView | null {
    const hf = hfEdit.value;
    if (hf?.headerFooter && opts.getHfPmView) {
      const v = opts.getHfPmView(hf.headerFooter);
      if (v) return v;
    }
    return opts.editorView.value;
  }

  function setPmSelection(anchor: number, head?: number) {
    const view = activeView();
    if (!view) return;
    try {
      const $anchor = view.state.doc.resolve(anchor);
      const $head = head !== undefined ? view.state.doc.resolve(head) : $anchor;
      const sel = TextSelection.between($anchor, $head);
      view.dispatch(view.state.tr.setSelection(sel));
    } catch {
      // Position invalid for this doc (e.g. body pos passed to HF view).
    }
  }

  function scrollVisiblePositionIntoView(pmPos: number) {
    scrollVisiblePositionIntoViewImpl(opts.pagesRef.value, opts.pagesViewportRef.value, pmPos);
  }

  function selectWord(pos: number) {
    selectWordImpl(opts.pagesRef.value, pos, setPmSelection, hfEdit.value?.position);
  }

  function selectParagraph(pos: number) {
    selectParagraphImpl(opts.pagesRef.value, pos, setPmSelection, hfEdit.value?.position);
  }

  function navigateToBookmark(bookmarkName: string) {
    const view = opts.editorView.value;
    if (!view) return;
    let targetPos: number | null = null;
    view.state.doc.descendants((node, pos) => {
      if (targetPos !== null) return false;
      const raw: unknown = node.attrs["bookmarks"];
      if (Array.isArray(raw)) {
        const bookmarks: readonly unknown[] = raw;
        for (const b of bookmarks) {
          if (b !== null && typeof b === "object" && "name" in b && b.name === bookmarkName) {
            targetPos = pos;
            return false;
          }
        }
      }
      return true;
    });
    if (targetPos === null) return;
    scrollVisiblePositionIntoView(targetPos);
    try {
      setPmSelection(Math.min(targetPos + 1, view.state.doc.content.size));
    } catch {
      // Bookmark target may be a non-text selectable position; fall back to the
      // start position so the click still moves the editor near the target.
      setPmSelection(targetPos);
    }
  }

  /**
   * Show / hide the "+" insert button as the cursor moves near a
   * table's edges. Hide is debounced through `TABLE_INSERT_HIDE_DELAY_MS`
   * so transient gaps between cells don't make the button flicker.
   */
  function handlePagesMouseMove(event: MouseEvent) {
    if (opts.readOnly.value) return;
    // Skip the hit-test during text drag-selects so the (+) doesn't
    // pop in mid-selection when the drag path crosses a table edge.
    if (isDragging) return;
    const pagesEl = opts.pagesRef.value;
    if (!pagesEl) return;
    const viewportEl = opts.pagesViewportRef.value;
    if (!viewportEl) return;
    if (!(event.target instanceof HTMLElement)) return;

    const hit = detectTableInsertHover({
      mouseX: event.clientX,
      mouseY: event.clientY,
      pagesContainer: pagesEl,
      target: event.target,
      hfEditMode: hfEdit.value?.position ?? null,
    });

    if (!hit) {
      if (tableInsertHideTimer === null) {
        tableInsertHideTimer = setTimeout(() => {
          tableInsertButton.value = null;
          tableInsertHideTimer = null;
        }, TABLE_INSERT_HIDE_DELAY_MS);
      }
      return;
    }

    // viewportEl carries `transform: scale(zoom)`; its rect is screen-space.
    // The button is an absolutely-positioned child of that scaled element, so
    // its left/top live in the element's own (unscaled) coords. Divide the
    // screen-space offset by zoom or it gets re-scaled and drifts (#928).
    const zoom = opts.zoom.value || 1;
    const viewportRect = viewportEl.getBoundingClientRect();
    tableInsertButton.value = {
      type: hit.type,
      x: (hit.clientX - viewportRect.left) / zoom,
      y: (hit.clientY - viewportRect.top) / zoom,
      cellPmPos: hit.cellPmPos,
    };
    clearTableInsertTimer();
  }

  /**
   * Insert a row below / column to the right of the target cell. The
   * core `addRowBelow` / `addColumnRight` commands read the current
   * PM selection to know which cell to extend, so we plant a caret
   * inside the hovered cell first.
   */
  function handleTableInsertClick(event: MouseEvent) {
    event.preventDefault();
    event.stopPropagation();
    const btn = tableInsertButton.value;
    const view = opts.editorView.value;
    if (!btn || !view) return;
    const tr = view.state.tr.setSelection(TextSelection.create(view.state.doc, btn.cellPmPos + 1));
    view.dispatch(tr);
    const cmds = opts.getCommands();
    const factory = btn.type === "row" ? cmds["addRowBelow"] : cmds["addColumnRight"];
    if (!factory) return;
    factory()(view.state, (t) => view.dispatch(t), view);
    tableInsertButton.value = null;
    view.focus();
  }

  /**
   * Single-click on a hyperlink → surface the popup or navigate internal
   * bookmarks. Browser default navigation stays suppressed so drag-selects
   * ending on links do not unexpectedly leave the document.
   */
  function handlePagesClick(event: MouseEvent) {
    const el = event.target;
    const anchor = el instanceof HTMLElement ? el.closest<HTMLAnchorElement>("a[href]") : null;
    if (!anchor) return;
    event.preventDefault();
    const href = anchor.getAttribute("href") ?? "";
    if (href.startsWith("#")) {
      const bookmarkName = href.slice(1);
      if (bookmarkName) navigateToBookmark(bookmarkName);
      return;
    }
    // Route through the active view (the header/footer PM when editing one),
    // not the body PM. Empty hrefs still surface the popup (so the user can
    // add/edit a URL), matching React.
    const view = activeView();
    const hasRangeSelection = view && view.state.selection.from !== view.state.selection.to;
    if (hasRangeSelection) return;
    // Compute popup position relative to the pages viewport so the popup
    // can render inside the scroll context — the browser then repositions
    // it on scroll via CSS alone, no JS listener needed.
    const viewport = opts.pagesViewportRef.value;
    if (!viewport) return;
    const vpRect = viewport.getBoundingClientRect();
    const linkRect = anchor.getBoundingClientRect();
    const title = anchor.getAttribute("title");
    opts.hyperlinkPopupData.value = {
      href,
      displayText: anchor.textContent ?? "",
      position: {
        top: linkRect.bottom - vpRect.top + viewport.scrollTop + 4,
        left: linkRect.left - vpRect.left + viewport.scrollLeft,
      },
      ...(title ? { tooltip: title } : {}),
    };
  }

  function handlePagesDoubleClick(event: MouseEvent) {
    if (!(event.target instanceof HTMLElement)) return;
    const target = event.target;
    const headerEl = target.closest<HTMLElement>(".layout-page-header");
    const footerEl = target.closest<HTMLElement>(".layout-page-footer");
    const hfEl = headerEl ?? footerEl;
    if (!hfEl) return;

    const position: "header" | "footer" = headerEl ? "header" : "footer";

    // No scroll-to-page-1 — HF content is shared across pages by `r:id`,
    // so edits propagate to every painted instance automatically.
    const doc = opts.getDocument();
    if (!doc?.package) return;

    // Resolve the HF for the current section.
    const sp =
      doc.package.document?.sections?.[0]?.properties ??
      doc.package.document?.finalSectionProperties ??
      null;
    const refs = position === "header" ? sp?.headerReferences : sp?.footerReferences;
    const map = position === "header" ? doc.package.headers : doc.package.footers;
    // Default ref takes priority; fall back to `first` if the doc only ships first.
    const refEntry =
      refs?.find((r) => r.type === "default") ?? refs?.find((r) => r.type === "first") ?? null;
    let rId: string | null = refEntry?.rId ?? null;
    let hf: HeaderFooter | null = rId ? (map?.get(rId) ?? null) : null;

    // Materialise an empty HF part if none exists for this section yet.
    // Without this, double-clicking an empty header is a no-op.
    if (!hf) {
      if (!sp) return;
      const hdrFtrType = "default" as const;
      const newRId = `rId_new_${position}_${hdrFtrType}`;
      const emptyHf: HeaderFooter = {
        type: position,
        hdrFtrType,
        content: [{ type: "paragraph", content: [] }],
      };
      const mapKey = position === "header" ? "headers" : "footers";
      const refKey = position === "header" ? "headerReferences" : "footerReferences";
      const newMap = new Map(doc.package[mapKey] ?? []);
      newMap.set(newRId, emptyHf);

      // Register a relationship so the serializer emits content types + doc rels.
      const existingRels = doc.package.relationships;
      const usedTargets = new Set<string>();
      for (const rel of existingRels?.values() ?? []) {
        if (rel.target) usedTargets.add(rel.target);
      }
      let targetNum = 1;
      while (usedTargets.has(`${position}${targetNum}.xml`)) targetNum++;
      const relType =
        position === "header"
          ? "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
          : "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
      const newRels = new Map(existingRels);
      newRels.set(newRId, {
        id: newRId,
        type: relType,
        target: `${position}${targetNum}.xml`,
      });

      // Create a fresh Document with the new HF wired in so any computeds
      // watching document identity refire and undo/redo can track the event.
      const newRef = { type: hdrFtrType, rId: newRId };
      const newSp: SectionProperties = { ...sp, [refKey]: [...(sp[refKey] ?? []), newRef] };
      const prevBody = doc.package.document;
      const mappedSections = prevBody?.sections?.map((s, i) =>
        i === 0 ? { ...s, properties: newSp } : s,
      );
      const newFinal =
        prevBody?.finalSectionProperties === sp ? newSp : prevBody?.finalSectionProperties;
      const newDoc: Document = {
        ...doc,
        package: {
          ...doc.package,
          [mapKey]: newMap,
          relationships: newRels,
          document: prevBody
            ? {
                ...prevBody,
                ...(mappedSections !== undefined && { sections: mappedSections }),
                ...(newFinal !== undefined && { finalSectionProperties: newFinal }),
              }
            : prevBody,
        },
      };
      rId = newRId;
      hf = emptyHf;
      opts.setDocument?.(newDoc);
      opts.syncHfPMs?.();
      opts.reLayout();
      opts.emit("change", newDoc);
    }

    // Bounding rect relative to the pages-viewport. zoom is applied via
    // CSS transform on the viewport, so use the unscaled element coords.
    const viewport = opts.pagesViewportRef.value;
    if (!viewport) return;
    const elRect = hfEl.getBoundingClientRect();
    const vpRect = viewport.getBoundingClientRect();
    const z = opts.zoom.value || 1;
    hfEdit.value = {
      position,
      rId,
      headerFooter: hf,
      targetRect: {
        top: (elRect.top - vpRect.top + viewport.scrollTop) / z,
        left: (elRect.left - vpRect.left + viewport.scrollLeft) / z,
        width: elRect.width / z,
        height: elRect.height / z,
      },
    };
  }

  function handleHfSave(content: BlockContent[]) {
    const doc = opts.getDocument();
    const edit = hfEdit.value;
    if (!doc?.package || !edit) return;
    const map = edit.position === "header" ? doc.package.headers : doc.package.footers;
    if (!map || !edit.rId) return;
    const existing = map.get(edit.rId);
    if (existing) {
      existing.content = content;
    }
    // After the inline overlay writes back into `pkg.headers/footers[rId]
    // .content`, re-sync re-mounts the persistent HF EditorView from the new
    // content so the painter sees the saved version.
    opts.syncHfPMs?.();
    opts.reLayout();
    opts.emit("change", doc);
  }

  function handleHfRemove() {
    const doc = opts.getDocument();
    const edit = hfEdit.value;
    if (!doc?.package || !edit || !edit.rId) {
      hfEdit.value = null;
      return;
    }
    // Actually remove the header/footer: drop the part from the headers/footers
    // map AND strip every section reference that points at it. Clearing
    // `content` alone left an empty header/footer still referenced by the
    // section, so it kept rendering.
    const mapKey = edit.position === "header" ? "headers" : "footers";
    const refKey = edit.position === "header" ? "headerReferences" : "footerReferences";
    const rId = edit.rId;

    const newMap = new Map(doc.package[mapKey] ?? []);
    newMap.delete(rId);

    // Strip the reference everywhere a section can carry it: each
    // sections[].properties, finalSectionProperties, and any mid-body section
    // break (a paragraph's pPr/sectPr in body.content).
    const stripRefs = (sp: SectionProperties): SectionProperties => ({
      ...sp,
      [refKey]: (sp[refKey] ?? []).filter((r) => r.rId !== rId),
    });
    const stripBlock = <T extends BlockContent>(block: T): T =>
      "sectionProperties" in block && block.sectionProperties
        ? { ...block, sectionProperties: stripRefs(block.sectionProperties) }
        : block;

    const body = doc.package.document;
    const strippedSections = body?.sections?.map((s) => ({
      ...s,
      properties: stripRefs(s.properties),
    }));
    const strippedFinal = body?.finalSectionProperties
      ? stripRefs(body.finalSectionProperties)
      : undefined;
    const newDoc: Document = {
      ...doc,
      package: {
        ...doc.package,
        [mapKey]: newMap,
        document: body
          ? {
              ...body,
              content: body.content.map(stripBlock),
              ...(strippedSections !== undefined && { sections: strippedSections }),
              ...(strippedFinal !== undefined && { finalSectionProperties: strippedFinal }),
            }
          : body,
      },
    };

    hfEdit.value = null;
    opts.setDocument?.(newDoc);
    opts.syncHfPMs?.();
    opts.reLayout();
    opts.emit("change", newDoc);
  }

  function handlePagesMouseDown(event: MouseEvent) {
    if (event.button !== 0) return;
    if (opts.imageInteracting.value) return;
    const body = opts.editorView.value;
    if (!body) return;

    const target = event.target;
    const targetEl = target instanceof HTMLElement ? target : null;

    // HF mode: clicks OUTSIDE the painted HF area close edit mode and refocus
    // the body PM. The body-PM-selection branch below also falls through, so
    // the next keystroke lands at the click site in the body.
    if (hfEdit.value && targetEl) {
      const isInHfArea =
        targetEl.closest(".layout-page-header") ||
        targetEl.closest(".layout-page-footer") ||
        targetEl.closest(".hf-editor");
      if (!isInHfArea) {
        hfEdit.value = null;
        body.focus();
        // Fall through — body-selection path resolves cursor at click coord.
      }
    }

    // Resolve the PM the user is currently editing (HF when active, body
    // otherwise). Every gesture below dispatches on this view.
    const view = activeView() ?? body;

    // Table resize: column / row / right-edge handles claim the gesture
    // regardless of which doc the cells belong to.
    if (!opts.readOnly.value && opts.tableResize.tryStartResize(event, view)) {
      return;
    }

    // Image click → NodeSelection on the active doc.
    const imageEl = findImageElement(target);
    if (imageEl) {
      event.preventDefault();
      event.stopPropagation();
      const pmStart = Number(imageEl.dataset["pmStart"]);
      if (!Number.isNaN(pmStart)) {
        try {
          view.dispatch(view.state.tr.setSelection(NodeSelection.create(view.state.doc, pmStart)));
        } catch {
          // Position may not be a valid node anchor.
        }
        opts.selectedImage.value = {
          element: imageEl,
          pmPos: pmStart,
          width: imageEl.offsetWidth,
          height: imageEl.offsetHeight,
        };
        opts.clearOverlay();
      }
      view.focus();
      return;
    }

    // Click outside an image clears the image selection.
    opts.selectedImage.value = null;

    event.preventDefault();

    const pos = resolvePos(event.clientX, event.clientY);
    if (pos === null) {
      view.focus();
      return;
    }

    // Multi-click detection
    const now = Date.now();
    if (now - lastClickTime < MULTI_CLICK_DELAY && lastClickPos === pos) {
      clickCount++;
    } else {
      clickCount = 1;
    }
    lastClickTime = now;
    lastClickPos = pos;

    if (clickCount === 2) {
      selectWord(pos);
    } else if (clickCount >= 3) {
      selectParagraph(pos);
      clickCount = 0;
    } else if (event.shiftKey) {
      // Single click — shift-click extends, plain click collapses.
      const { from } = view.state.selection;
      setPmSelection(from, pos);
    } else {
      setPmSelection(pos);
      dragAnchor = pos;
      isDragging = true;
      // Record the cell under the press so a drag across cells promotes to a
      // CellSelection (null when the press isn't inside a table).
      cellDrag.begin(findCellPosFromPmPos(view, pos));
    }

    view.focus();
  }

  function handleMouseMove(event: MouseEvent) {
    if (!isDragging || dragAnchor === null) return;
    const pos = resolvePos(event.clientX, event.clientY);
    if (pos !== null) {
      const view = activeView();
      // A drag that crosses cell boundaries becomes a CellSelection; when it
      // does, skip the text-selection update for this move.
      if (view && cellDrag.update(view, pos, event.clientX)) {
        dragAutoScroll.updateMousePosition(event.clientX, event.clientY);
        return;
      }
      if (pos !== dragAnchor) {
        setPmSelection(dragAnchor, pos);
      }
    }
    // Drive edge auto-scroll while dragging.
    dragAutoScroll.updateMousePosition(event.clientX, event.clientY);
  }

  function handleMouseUp() {
    isDragging = false;
    cellDrag.end();
    dragAutoScroll.stopAutoScroll();
  }

  function handleViewportScroll() {
    const container = opts.pagesViewportRef.value;
    const lay = opts.layout.value;
    if (!container || !lay || lay.pages.length === 0) return;

    const scrollTop = container.scrollTop;
    const totalPages = lay.pages.length;
    const PAGE_GAP = 24; // matches DEFAULT_PAGE_GAP in useDocxEditor
    const PADDING_TOP = 24;

    const viewportCenter = scrollTop + container.clientHeight / 2;
    let accumulatedY = PADDING_TOP;
    let currentPage = 1;
    for (let i = 0; i < lay.pages.length; i++) {
      const page = lay.pages[i];
      if (!page) continue;
      const pageEnd = accumulatedY + page.size.h;
      if (viewportCenter < pageEnd) {
        currentPage = i + 1;
        break;
      }
      accumulatedY = pageEnd + PAGE_GAP;
      currentPage = i + 2;
    }
    currentPage = Math.min(currentPage, totalPages);

    scrollPageInfo.value = { currentPage, totalPages, visible: true };

    if (scrollFadeTimer) clearTimeout(scrollFadeTimer);
    scrollFadeTimer = setTimeout(() => {
      scrollPageInfo.value = { ...scrollPageInfo.value, visible: false };
    }, 600);
  }

  onMounted(() => {
    window.addEventListener("mousemove", handleMouseMove);
    window.addEventListener("mouseup", handleMouseUp);
    opts.pagesViewportRef.value?.addEventListener("scroll", handleViewportScroll, {
      passive: true,
    });
  });

  onBeforeUnmount(() => {
    clearTableInsertTimer();
    window.removeEventListener("mousemove", handleMouseMove);
    window.removeEventListener("mouseup", handleMouseUp);
    opts.pagesViewportRef.value?.removeEventListener("scroll", handleViewportScroll);
    if (scrollFadeTimer) clearTimeout(scrollFadeTimer);
  });

  return {
    tableInsertButton,
    hfEdit,
    scrollPageInfo,
    resolvePos,
    setPmSelection,
    scrollVisiblePositionIntoView,
    navigateToBookmark,
    handlePagesMouseDown,
    handlePagesMouseMove,
    handlePagesClick,
    handlePagesDoubleClick,
    handleTableInsertClick,
    clearTableInsertTimer,
    handleHfSave,
    handleHfRemove,
  };
}
