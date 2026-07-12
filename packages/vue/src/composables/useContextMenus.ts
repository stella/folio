/**
 * Right-click menu composable — owns the `contextMenu` /
 * `imageContextMenu` state plus the trio of handlers that surface them
 * (`handleContextMenu`, `handleSelectedImageContextMenu`,
 * `handleContextMenuAction`) and the wrap-select bridge used by the
 * image menu. Reads `selectedImage` from `useImageActions` and writes
 * back into it for cut/delete-image actions. Pages-level selection
 * primitives (`clearOverlay`, `setPmSelection`, `resolvePos`) are
 * passed in as callbacks so this composable can live independently of
 * the still-in-parent pointer wiring.
 */

import { computed, ref, type ComputedRef, type Ref, type ShallowRef, type VNodeChild } from "vue";
import { TextSelection, NodeSelection, type Command } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import {
  captureInlinePositionEmu,
  findImageElement,
  type ImageLayoutTarget,
} from "@stll/folio-core/layout-painter/imageLayout";
import { getTableContext } from "@stll/folio-core/prosemirror/extensions/nodes/TableExtension";
import type { WrapType } from "@stll/folio-core/docx/wrapTypes";
import {
  copyImageToClipboard,
  pasteFromClipboard,
  triggerReplaceImage,
} from "../utils/imageClipboard";
import type { ImageSelectionInfo } from "../components/imageSelectionTypes";
import type {
  ImageContextMenuState,
  ImageContextMenuTextAction,
} from "../components/imageContextMenuTypes";
import { useTranslation } from "../i18n";

type CommandFactory = (...args: readonly unknown[]) => Command;

type CssFloat = "left" | "right" | "none" | null;

const WRAP_TYPE_SET: ReadonlySet<string> = new Set<WrapType>([
  "inline",
  "square",
  "tight",
  "through",
  "topAndBottom",
  "behind",
  "inFront",
]);

function isWrapTypeValue(value: unknown): value is WrapType {
  return typeof value === "string" && WRAP_TYPE_SET.has(value);
}

function readWrapType(value: unknown): WrapType {
  return isWrapTypeValue(value) ? value : "inline";
}

function readCssFloat(value: unknown): CssFloat {
  return value === "left" || value === "right" || value === "none" ? value : null;
}

export type TextContextMenuState = {
  isOpen: boolean;
  position: { x: number; y: number };
  hasSelection: boolean;
  inTable: boolean;
  onImage: boolean;
  canMergeCells: boolean;
  canSplitCell: boolean;
};

/**
 * One host-provided context-menu entry (the `customContextMenuItems` prop).
 * Structural mirror of React's `DocxEditorProps.customContextMenuItems` items;
 * `icon` is a Vue `VNodeChild` rather than a React node.
 */
export type CustomContextMenuItem = {
  id: string;
  label: string;
  requiresSelection?: boolean;
  icon?: VNodeChild;
};

export type UseContextMenusOptions = {
  editorView: Ref<EditorView | null>;
  selectedImage: ShallowRef<ImageSelectionInfo | null>;
  zoom: Ref<number>;
  showImageProperties: Ref<boolean>;
  getCommands: () => Record<string, CommandFactory>;
  clearOverlay: () => void;
  setPmSelection: (anchor: number, head?: number) => void;
  resolvePos: (clientX: number, clientY: number) => number | null;
  /** Host-injected menu entries, appended after the built-ins (lead the menu). */
  customContextMenuItems?: () => readonly CustomContextMenuItem[] | undefined;
  /**
   * Fires with the item id (menu-namespace stripped) and the selection range
   * captured when the menu opened. Right-click can collapse the live PM
   * selection, so the captured range is preferred when non-empty.
   */
  onCustomContextAction?: (id: string, selectionRange: { from: number; to: number }) => void;
};

export type UseContextMenusReturn = {
  contextMenu: Ref<TextContextMenuState>;
  imageContextMenu: Ref<ImageContextMenuState | null>;
  imageContextMenuTextActions: ComputedRef<ImageContextMenuTextAction[]>;
  /** Host items to lead the text menu, already filtered by `requiresSelection`. */
  customTextMenuItems: ComputedRef<CustomContextMenuItem[]>;
  handleContextMenu: (event: MouseEvent) => void;
  handleSelectedImageContextMenu: (event: MouseEvent) => void;
  handleImageWrapSelect: (target: ImageLayoutTarget) => void;
  handleContextMenuAction: (action: string) => void;
};

export function useContextMenus(opts: UseContextMenusOptions): UseContextMenusReturn {
  const { t } = useTranslation();

  const contextMenu = ref<TextContextMenuState>({
    isOpen: false,
    position: { x: 0, y: 0 },
    hasSelection: false,
    inTable: false,
    onImage: false,
    canMergeCells: false,
    canSplitCell: false,
  });

  // Image-specific right-click menu — shows wrap-mode options instead of the
  // generic text menu when the user right-clicks a rendered image.
  const imageContextMenu = ref<ImageContextMenuState | null>(null);

  // Selection range captured when the text menu opens. Right-click can collapse
  // the live PM selection, so `custom:` actions resolve against this snapshot
  // (preferring it when non-empty) rather than the post-click selection.
  const capturedSelectionRange = ref<{ from: number; to: number }>({ from: 0, to: 0 });

  // Host menu entries to lead the text menu, filtered by `requiresSelection`
  // against the current selection (mirrors React's DocxEditor menu build).
  const customTextMenuItems = computed<CustomContextMenuItem[]>(() =>
    (opts.customContextMenuItems?.() ?? []).filter(
      (item) => !item.requiresSelection || contextMenu.value.hasSelection,
    ),
  );

  // Cut / Copy / Paste / Delete items appended below the layout choices in
  // the image context menu so users don't have to flip menus to do
  // clipboard work on a selected image.
  const imageContextMenuTextActions = computed<ImageContextMenuTextAction[]>(() => [
    { action: "cut", label: t("contextMenu.cut"), shortcut: t("contextMenu.cutShortcut") },
    { action: "copy", label: t("contextMenu.copy"), shortcut: t("contextMenu.copyShortcut") },
    {
      action: "paste",
      label: t("contextMenu.paste"),
      shortcut: t("contextMenu.pasteShortcut"),
      dividerAfter: true,
    },
    {
      action: "delete",
      label: t("contextMenu.delete"),
      shortcut: t("contextMenu.deleteShortcut"),
    },
  ]);

  function runCommand(view: EditorView, command: Command): void {
    command(view.state, (tr) => view.dispatch(tr), view);
  }

  function buildImageMenuState(
    imageEl: HTMLElement,
    pmPos: number,
    position: { x: number; y: number },
    wrapType: WrapType,
    cssFloat: CssFloat,
  ): ImageContextMenuState {
    const inlinePositionEmu =
      wrapType === "inline" ? captureInlinePositionEmu(imageEl, opts.zoom.value) : undefined;
    return {
      open: true,
      position,
      pmPos,
      currentWrapType: wrapType,
      currentCssFloat: cssFloat,
      ...(inlinePositionEmu !== undefined && { inlinePositionEmu }),
    };
  }

  function handleContextMenu(event: MouseEvent) {
    const view = opts.editorView.value;
    if (!view) return;

    // Check if right-click is on an image
    const imageEl = findImageElement(event.target);
    if (imageEl) {
      const pmStart = Number(imageEl.dataset["pmStart"]);
      if (!Number.isNaN(pmStart)) {
        try {
          const sel = NodeSelection.create(view.state.doc, pmStart);
          view.dispatch(view.state.tr.setSelection(sel));
        } catch {
          /* ignore */
        }
        opts.selectedImage.value = {
          element: imageEl,
          pmPos: pmStart,
          width: imageEl.offsetWidth,
          height: imageEl.offsetHeight,
        };
        opts.clearOverlay();

        // Image right-click takes priority over the text context menu —
        // surface the layout-options menu instead and bail out.
        const node = view.state.doc.nodeAt(pmStart);
        if (node && node.type.name === "image") {
          imageContextMenu.value = buildImageMenuState(
            imageEl,
            pmStart,
            { x: event.clientX, y: event.clientY },
            readWrapType(node.attrs["wrapType"]),
            readCssFloat(node.attrs["cssFloat"]),
          );
          contextMenu.value.isOpen = false;
          return;
        }
      }
    }

    // Move the PM caret to the right-click point unless the click landed
    // inside the current selection (or exactly on a collapsed caret —
    // re-dispatching the same position would force a needless re-layout).
    // Mirrors React's PagedEditor: table ops and other caret-scoped
    // actions then operate on the cell/run the user actually clicked.
    {
      const { from, to } = view.state.selection;
      const clickPos = opts.resolvePos(event.clientX, event.clientY);
      if (clickPos !== null && (clickPos < from || clickPos > to)) {
        try {
          opts.setPmSelection(clickPos);
        } catch {
          // resolved position may be out of range after a concurrent edit
        }
      }
    }

    const tableCtx = getTableContext(view.state);
    const { empty, from, to } = view.state.selection;

    // Snapshot the (post-caret-move) selection so a host `custom:` action can
    // recover the range even if the click collapsed it.
    capturedSelectionRange.value = { from, to };

    // Right-clicking outside an image clears any open image context menu
    // — otherwise the image menu can stay visible while TextContextMenu
    // is shown over a different element. Mirrors React's PagedEditor
    // exclusivity (only one of the two menus visible at a time).
    if (imageContextMenu.value) imageContextMenu.value = null;

    contextMenu.value = {
      isOpen: true,
      position: { x: event.clientX, y: event.clientY },
      hasSelection: !empty,
      inTable: tableCtx.isInTable,
      onImage: !!imageEl,
      canMergeCells: !!tableCtx.hasMultiCellSelection,
      canSplitCell: !!tableCtx.canSplitCell,
    };
  }

  function handleSelectedImageContextMenu(event: MouseEvent) {
    const view = opts.editorView.value;
    const sel = opts.selectedImage.value;
    if (!view || !sel) return;
    const node = view.state.doc.nodeAt(sel.pmPos);
    if (!node || node.type.name !== "image") return;
    imageContextMenu.value = buildImageMenuState(
      sel.element,
      sel.pmPos,
      { x: event.clientX, y: event.clientY },
      readWrapType(node.attrs["wrapType"]),
      readCssFloat(node.attrs["cssFloat"]),
    );
    contextMenu.value.isOpen = false;
  }

  function handleImageWrapSelect(target: ImageLayoutTarget) {
    const view = opts.editorView.value;
    const state = imageContextMenu.value;
    if (!view || !state) return;
    const optsArg =
      state.inlinePositionEmu && target !== "inline"
        ? { initialPositionEmu: state.inlinePositionEmu }
        : undefined;
    const factory = opts.getCommands()["setImageWrapType"];
    if (!factory) return;
    runCommand(view, factory(state.pmPos, target, optsArg));
    view.focus();
  }

  function handleContextMenuAction(action: string) {
    const view = opts.editorView.value;
    if (!view) return;
    const cmds = opts.getCommands();

    // Host-injected entries carry a `custom:` namespace. Resolve the range from
    // the captured snapshot (preferred when non-empty), otherwise the live
    // selection — mirrors React's DocxEditor `custom:` branch.
    if (action.startsWith("custom:")) {
      const captured = capturedSelectionRange.value;
      const range =
        captured.from !== captured.to
          ? captured
          : { from: view.state.selection.from, to: view.state.selection.to };
      opts.onCustomContextAction?.(action.slice("custom:".length), range);
      view.focus();
      return;
    }

    switch (action) {
      case "cut":
        if (opts.selectedImage.value) {
          const pos = opts.selectedImage.value.pmPos;
          copyImageToClipboard(view, pos);
          const node = view.state.doc.nodeAt(pos);
          if (node) {
            view.dispatch(view.state.tr.delete(pos, pos + node.nodeSize));
            opts.selectedImage.value = null;
          }
        } else {
          // Focus the hidden PM first so the browser's clipboard op targets it.
          // The hidden editor is off-screen (left: -9999px) and not focused
          // after a context-menu click, so execCommand would otherwise no-op
          // (#929). The trailing view.focus() below runs too late.
          view.focus();
          document.execCommand("cut");
        }
        break;
      case "copy":
        if (opts.selectedImage.value) {
          copyImageToClipboard(view, opts.selectedImage.value.pmPos);
        } else {
          view.focus();
          document.execCommand("copy");
        }
        break;
      case "paste":
        void pasteFromClipboard(view);
        break;
      case "pasteAsPlainText":
        // Strip all formatting — insert the clipboard's text/plain only.
        navigator.clipboard
          .readText()
          .then((text) => {
            if (text) view.dispatch(view.state.tr.insertText(text).scrollIntoView());
            return undefined;
          })
          .catch(() => {
            // Clipboard read denied — nothing to paste.
            return undefined;
          });
        break;
      case "delete": {
        const { from, to } = view.state.selection;
        if (from !== to) view.dispatch(view.state.tr.delete(from, to));
        break;
      }
      case "selectAll": {
        const sel = TextSelection.create(view.state.doc, 0, view.state.doc.content.size);
        view.dispatch(view.state.tr.setSelection(sel));
        break;
      }
      case "imageProperties":
        if (opts.selectedImage.value) {
          opts.showImageProperties.value = true;
        }
        break;
      case "replaceImage":
        if (opts.selectedImage.value) {
          triggerReplaceImage(view, opts.selectedImage.value.pmPos);
        }
        break;
      case "deleteImage": {
        if (opts.selectedImage.value) {
          const pos = opts.selectedImage.value.pmPos;
          const node = view.state.doc.nodeAt(pos);
          if (node) {
            view.dispatch(view.state.tr.delete(pos, pos + node.nodeSize));
            opts.selectedImage.value = null;
          }
        }
        break;
      }
      case "addRowAbove":
      case "addRowBelow":
      case "deleteRow":
      case "addColumnLeft":
      case "addColumnRight":
      case "deleteColumn":
      case "mergeCells":
      case "splitCell":
      case "selectTable":
      case "deleteTable": {
        const factory = cmds[action];
        if (factory) runCommand(view, factory());
        break;
      }
    }
    view.focus();
  }

  return {
    contextMenu,
    imageContextMenu,
    imageContextMenuTextActions,
    customTextMenuItems,
    handleContextMenu,
    handleSelectedImageContextMenu,
    handleImageWrapSelect,
    handleContextMenuAction,
  };
}
