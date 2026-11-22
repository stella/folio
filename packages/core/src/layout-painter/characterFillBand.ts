import type { MeasuredLine } from "../layout-engine/types";
import { COMMENT_ANCHOR_SELECTOR } from "../render-dom/commentAnchorAttributes";
import {
  AUTHORED_BACKGROUND_COLOR_VAR,
  setAuthoredBackgroundColor,
} from "./documentColors";

type PendingFill = {
  line: HTMLElement;
  band: NonNullable<MeasuredLine["characterFillBand"]>;
  layer: HTMLElement;
  runs: { element: HTMLElement; color: string }[];
};

export const CHARACTER_FILL_BAND_ATTRIBUTE = "data-character-fill-band";
export const CHARACTER_FILL_BACKINGS_ATTRIBUTE = "data-character-fill-backings";

const pendingFills: PendingFill[] = [];

function measureFill({ line, runs }: PendingFill) {
  if (!line.isConnected) return undefined;
  const lineRect = line.getBoundingClientRect();
  if (lineRect.width <= 0) return undefined;
  return runs.map(({ element, color }) => ({
    element,
    color,
    rects: Array.from(element.getClientRects())
      .filter((rect) => rect.width > 0)
      .map((rect) => ({
        leftPercent: ((rect.left - lineRect.left) / lineRect.width) * 100,
        widthPercent: (rect.width / lineRect.width) * 100,
      })),
  }));
}

function paintPendingFills(): void {
  // Read every mounted run before writing. The browser owns inline advances,
  // including shaping, justified spaces and bidi order.
  const measured = pendingFills.splice(0).map((fill) => ({
    fill,
    runs: measureFill(fill),
  }));
  for (const { fill, runs } of measured) {
    const { line, band, layer } = fill;
    if (!runs) {
      // A standalone paint may never attach the line to a document.
      for (const { element, color } of fill.runs)
        setAuthoredBackgroundColor(element.style, color);
      continue;
    }
    for (const { element, color, rects } of runs) {
      for (const { leftPercent, widthPercent } of rects) {
        const backing = line.ownerDocument.createElement("span");
        backing.style.position = "absolute";
        backing.style.left = `${leftPercent}%`;
        backing.style.width = `${widthPercent}%`;
        backing.style.top = `${band.topPx}px`;
        backing.style.height = `${band.heightPx}px`;
        backing.style.opacity = element.style.opacity;
        setAuthoredBackgroundColor(backing.style, color);
        layer.append(backing);
      }
    }
  }
}

export function scheduleCharacterFillBand(
  line: HTMLElement,
  band: PendingFill["band"],
  textClass: string,
): void {
  if (band.heightPx <= 0) return;
  const runSelector = `.${textClass}[style*="${AUTHORED_BACKGROUND_COLOR_VAR}"]:not(${COMMENT_ANCHOR_SELECTOR})`;
  const runs = Array.from(
    line.querySelectorAll<HTMLElement>(runSelector),
  ).flatMap((element) => {
    const color = element.style.getPropertyValue(AUTHORED_BACKGROUND_COLOR_VAR);
    if (!color) return [];
    element.setAttribute("data-character-fill-color", color);
    element.style.removeProperty("background-color");
    element.style.removeProperty(AUTHORED_BACKGROUND_COLOR_VAR);
    return [{ element, color }];
  });
  if (runs.length === 0) return;

  line.setAttribute(
    CHARACTER_FILL_BAND_ATTRIBUTE,
    `${band.topPx}:${band.heightPx}`,
  );
  line.style.position = "relative";
  line.style.isolation = "isolate";
  const layer = line.ownerDocument.createElement("span");
  layer.setAttribute(CHARACTER_FILL_BACKINGS_ATTRIBUTE, "");
  layer.setAttribute("aria-hidden", "true");
  layer.contentEditable = "false";
  layer.style.position = "absolute";
  layer.style.inset = "0";
  layer.style.pointerEvents = "none";
  layer.style.zIndex = "-1";
  line.append(layer);

  if (pendingFills.length === 0) queueMicrotask(paintPendingFills);
  pendingFills.push({ line, band, layer, runs });
}
