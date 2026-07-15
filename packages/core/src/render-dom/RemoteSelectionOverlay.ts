import type { HiddenProseMirrorRemoteSelection } from "../controller/hiddenEditorManager";
import { createRenderedDomContext } from "./RenderedDomContext";
import type {
  RenderedDomContext,
  RenderedDomPoint,
  RenderedDomRect,
} from "./RenderedDomContext";

const CARET_CLASS = "folio-remote-selection-caret";
const LABEL_CLASS = "folio-remote-selection-label";
const RANGE_CLASS = "folio-remote-selection-rect";
const OWNED_OVERLAY_SELECTOR = `.${CARET_CLASS}, .${LABEL_CLASS}, .${RANGE_CLASS}`;

export type SyncRemoteSelectionOverlayOptions = {
  pagesContainer: HTMLElement;
  selections: readonly HiddenProseMirrorRemoteSelection[];
  zoom: number;
  zIndex?: number;
  renderedDomContext?: RenderedDomContext;
};

export type RemoteSelectionOverlayGeometry = {
  caret: RenderedDomPoint | null;
  rects: RenderedDomRect[];
  selection: HiddenProseMirrorRemoteSelection;
};

export const resolveRemoteSelectionOverlayGeometry = (
  renderedDomContext: RenderedDomContext,
  selections: readonly HiddenProseMirrorRemoteSelection[],
): RemoteSelectionOverlayGeometry[] =>
  selections.map((selection) => ({
    caret: renderedDomContext.getCoordinatesForPosition(selection.head),
    rects: renderedDomContext.getRectsForRange(
      Math.min(selection.anchor, selection.head),
      Math.max(selection.anchor, selection.head),
    ),
    selection,
  }));

const setBox = (element: HTMLElement, rect: RenderedDomRect): void => {
  element.style.left = `${rect.x}px`;
  element.style.top = `${rect.y}px`;
  element.style.width = `${rect.width}px`;
  element.style.height = `${rect.height}px`;
};

const createOverlayElement = (
  pagesContainer: HTMLElement,
  className: string,
  zIndex: number,
): HTMLDivElement => {
  const element = pagesContainer.ownerDocument.createElement("div");
  element.className = className;
  element.style.position = "absolute";
  element.style.pointerEvents = "none";
  element.style.zIndex = String(zIndex);
  return element;
};

/** Paint collaborative cursors and selections over the rendered page DOM. */
export class RemoteSelectionOverlay {
  clear(pagesContainer: HTMLElement): void {
    for (const element of pagesContainer.querySelectorAll(OWNED_OVERLAY_SELECTOR)) {
      element.remove();
    }
  }

  sync({
    pagesContainer,
    selections,
    zoom,
    zIndex = 20,
    renderedDomContext = createRenderedDomContext(pagesContainer, zoom),
  }: SyncRemoteSelectionOverlayOptions): void {
    this.clear(pagesContainer);

    for (const { caret, rects, selection } of resolveRemoteSelectionOverlayGeometry(
      renderedDomContext,
      selections,
    )) {
      for (const rect of rects) {
        const element = createOverlayElement(pagesContainer, RANGE_CLASS, zIndex);
        element.dataset["clientId"] = String(selection.clientId);
        element.style.background = `color-mix(in srgb, ${selection.color} 24%, transparent)`;
        setBox(element, rect);
        pagesContainer.appendChild(element);
      }

      if (!caret) {
        continue;
      }

      const caretElement = createOverlayElement(pagesContainer, CARET_CLASS, zIndex + 1);
      caretElement.dataset["clientId"] = String(selection.clientId);
      caretElement.style.background = selection.color;
      setBox(caretElement, { ...caret, width: 2 });
      pagesContainer.appendChild(caretElement);

      const label = createOverlayElement(pagesContainer, LABEL_CLASS, zIndex + 2);
      label.dataset["clientId"] = String(selection.clientId);
      label.textContent = selection.name;
      label.style.background = selection.color;
      label.style.color = "var(--background, #fff)";
      label.style.fontSize = "10px";
      label.style.lineHeight = "1";
      label.style.padding = "2px 4px";
      label.style.borderRadius = "2px";
      label.style.left = `${caret.x}px`;
      label.style.top = `${Math.max(0, caret.y - 18)}px`;
      label.style.whiteSpace = "nowrap";
      pagesContainer.appendChild(label);
    }
  }
}
