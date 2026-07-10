/**
 * Image-actions composable — owns the `selectedImage` /
 * `imageInteracting` refs plus the toolbar/menu handlers that mutate a
 * selected image's wrap type or transform. (Inserting a fresh image is the
 * shared core `insertImageFromFile` flow, wired in `DocxEditor.vue`.)
 * Consumed downstream by `useContextMenus`, `usePagesPointer`,
 * and the selection-overlay update in the parent (which writes back
 * into `selectedImage` when the PM doc holds a NodeSelection on an
 * image). Does NOT own the right-click menus — those live in
 * `useContextMenus`.
 */

import { computed, ref, shallowRef, type ComputedRef, type Ref, type ShallowRef } from "vue";
import type { Command } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import {
  captureInlinePositionEmu,
  toolbarValueToLayoutTarget,
} from "@stll/folio-core/layout-painter/imageLayout";
import type { ImageSelectionInfo } from "../components/imageSelectionTypes";

type CommandFactory = (...args: readonly unknown[]) => Command;

export type UseImageActionsOptions = {
  editorView: Ref<EditorView | null>;
  zoom: Ref<number>;
  stateTick: Ref<number>;
  getCommands: () => Record<string, CommandFactory>;
};

export type ImageToolbarContext = {
  wrapType: string;
  displayMode: string;
  cssFloat: string | null;
};

export type UseImageActionsReturn = {
  selectedImage: ShallowRef<ImageSelectionInfo | null>;
  imageInteracting: Ref<boolean>;
  imageToolbarContext: ComputedRef<ImageToolbarContext | null>;
  handleToolbarImageWrap: (value: string) => void;
  handleImageTransform: (action: "rotateCW" | "rotateCCW" | "flipH" | "flipV") => void;
};

export function useImageActions(opts: UseImageActionsOptions): UseImageActionsReturn {
  // shallowRef so the wrapped HTMLElement isn't proxied — identity comparisons
  // downstream (ImageSelectionOverlay) rely on raw element references.
  const selectedImage: ShallowRef<ImageSelectionInfo | null> = shallowRef(null);

  // True while the overlay is mid-resize / move / rotate — gates the pages
  // mousedown handler so an in-flight image gesture isn't clobbered by a stray
  // click (mirrors React's PagedEditor.isImageInteractingRef).
  const imageInteracting = ref(false);

  // Toolbar image group: read the live image attrs from the PM doc at the
  // selected image's position so the wrap dropdown highlights the correct
  // active option. Only the three fields the toolbar dropdown reads — wrap
  // dropdown is the only UI element wired to this context in v1.
  const imageToolbarContext = computed<ImageToolbarContext | null>(() => {
    void opts.stateTick.value;
    const view = opts.editorView.value;
    const sel = selectedImage.value;
    if (!view || !sel) return null;
    const node = view.state.doc.nodeAt(sel.pmPos);
    if (!node || node.type.name !== "image") return null;
    const wrapRaw: unknown = node.attrs["wrapType"];
    const displayRaw: unknown = node.attrs["displayMode"];
    const cssFloatRaw: unknown = node.attrs["cssFloat"];
    return {
      wrapType: typeof wrapRaw === "string" ? wrapRaw : "inline",
      displayMode: typeof displayRaw === "string" ? displayRaw : "inline",
      cssFloat: typeof cssFloatRaw === "string" ? cssFloatRaw : null,
    };
  });

  function runCommand(view: EditorView, command: Command): void {
    command(view.state, (tr) => view.dispatch(tr), view);
  }

  // Toolbar wrap dropdown → core PM command. Translates the legacy
  // toolbar vocabulary via `toolbarValueToLayoutTarget` so this path
  // shares `setImageWrapType` with the right-click menu.
  function handleToolbarImageWrap(value: string) {
    const view = opts.editorView.value;
    const sel = selectedImage.value;
    if (!view || !sel) return;
    const target = toolbarValueToLayoutTarget(value);
    if (!target) return;
    const node = view.state.doc.nodeAt(sel.pmPos);
    const currentWrap: unknown = node?.attrs["wrapType"];
    const optsArg =
      currentWrap === "inline" && target !== "inline"
        ? { initialPositionEmu: captureInlinePositionEmu(sel.element, opts.zoom.value) }
        : undefined;
    const factory = opts.getCommands()["setImageWrapType"];
    if (!factory) return;
    runCommand(view, factory(sel.pmPos, target, optsArg));
    view.focus();
  }

  // Toolbar transform dropdown → mutate the selected image's
  // `transform` attribute. Rotate is folded mod 360, flip toggles bit
  // flags, then the parts are joined back into a CSS transform string.
  function handleImageTransform(action: "rotateCW" | "rotateCCW" | "flipH" | "flipV") {
    const view = opts.editorView.value;
    const sel = selectedImage.value;
    if (!view || !sel) return;
    const node = view.state.doc.nodeAt(sel.pmPos);
    if (!node || node.type.name !== "image") return;

    const transformRaw: unknown = node.attrs["transform"];
    const current = typeof transformRaw === "string" ? transformRaw : "";
    const degMatch = /rotate\((?<deg>-?\d+(?:\.\d+)?)deg\)/u.exec(current)?.groups?.["deg"];
    let rotation = degMatch ? Number.parseFloat(degMatch) : 0;
    let flipH = /scaleX\(-1\)/u.test(current);
    let flipV = /scaleY\(-1\)/u.test(current);

    if (action === "rotateCW") rotation = (rotation + 90) % 360;
    else if (action === "rotateCCW") rotation = (rotation - 90 + 360) % 360;
    else if (action === "flipH") flipH = !flipH;
    else if (action === "flipV") flipV = !flipV;

    const parts: string[] = [];
    if (rotation !== 0) parts.push(`rotate(${rotation}deg)`);
    if (flipH) parts.push("scaleX(-1)");
    if (flipV) parts.push("scaleY(-1)");
    const next = parts.length > 0 ? parts.join(" ") : null;

    const tr = view.state.tr.setNodeMarkup(sel.pmPos, undefined, {
      ...node.attrs,
      transform: next,
    });
    view.dispatch(tr.scrollIntoView());
    view.focus();
  }

  return {
    selectedImage,
    imageInteracting,
    imageToolbarContext,
    handleToolbarImageWrap,
    handleImageTransform,
  };
}
