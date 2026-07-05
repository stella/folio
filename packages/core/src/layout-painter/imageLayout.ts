/**
 * Image layout helpers shared between framework adapters (React, Vue, ...).
 *
 * Everything here is framework-agnostic: pure DOM math + pure functions over
 * OOXML wrap-type vocabulary. The corresponding UI bindings (right-click menu,
 * toolbar dropdown) live in each framework adapter and call into these.
 *
 * Ported from eigenpal/docx-editor `layout-painter/imageLayout.ts`.
 */

import type { WrapType } from "../docx/wrapTypes";
import type { FlowBlock, Measure } from "../layout-engine/types";
import type { ImageAttrs } from "../prosemirror/schema/nodes";
import { pixelsToEmu } from "../utils/units";
import type { BlockLookup } from "./index";

/**
 * Anchored wrap-type targets — the OOXML wrap types except `inline`.
 * Defined locally because the fork's `ImageExtension` does not export the
 * layout-target vocabulary the way upstream does.
 */
export type AnchorWrapType = Exclude<WrapType, "inline">;

/**
 * User-facing layout choices that mirror Word's Wrap Text menu. `squareLeft`
 * floats the image left (text wraps on the right); `squareRight` floats it
 * right; `inline` flows it as a glyph; the remaining targets are 1:1 with
 * their OOXML wrap types.
 */
export type ImageLayoutTarget = AnchorWrapType | "squareLeft" | "squareRight" | "inline";

// ============================================================================
// Class names emitted by the layout painter — single source of truth so hosts
// can hit-test rendered images without hard-coding selectors.
// ============================================================================

export const LAYOUT_IMAGE_CLASSES = {
  /** Inline image rendered inside `.layout-line`. */
  runImage: "layout-run-image",
  /** Block (centered, topAndBottom) image. */
  blockImage: "layout-block-image",
  /** Anchored image rendered in the page-level floating layer. */
  pageFloatingImage: "layout-page-floating-image",
  /** Anchored image rendered inside a table cell's floating layer. */
  cellFloatingImage: "layout-cell-floating-image",
  pageContent: "layout-page-content",
  paragraph: "layout-paragraph",
} as const;

const IMAGE_HIT_SELECTOR = [
  `.${LAYOUT_IMAGE_CLASSES.pageFloatingImage}`,
  `.${LAYOUT_IMAGE_CLASSES.cellFloatingImage}`,
  `.${LAYOUT_IMAGE_CLASSES.blockImage}`,
  `.${LAYOUT_IMAGE_CLASSES.runImage}`,
].join(", ");

// ============================================================================
// Hit-test
// ============================================================================

/**
 * Walk up from a click target to the nearest rendered image element, returning
 * just the element (no PM position parsing). Used by the left-click selection
 * path in both adapters — `data-pm-start` is read separately by callers that
 * need the position. Returns the element when the click was directly on an
 * inline `<img.layout-run-image>` OR inside one of the container classes
 * registered in `LAYOUT_IMAGE_CLASSES`.
 */
export function findImageElement(target: EventTarget | null): HTMLElement | null {
  if (!(target instanceof HTMLElement)) {
    return null;
  }
  // Direct hit on an inline `<img class="layout-run-image">` — return it as-is
  // since the inline image carries its own `data-pm-start` and behaves as the
  // selection target.
  if (target.tagName === "IMG" && target.classList.contains(LAYOUT_IMAGE_CLASSES.runImage)) {
    return target;
  }
  // Otherwise walk up to the nearest image container with `data-pm-start`.
  const container = target.closest<HTMLElement>(IMAGE_HIT_SELECTOR);
  if (container && container.dataset["pmStart"]) {
    return container;
  }
  return null;
}

// ============================================================================
// Position capture for inline → anchor transitions
// ============================================================================

/**
 * Capture the rendered position of an inline image as EMUs, normalised to
 * unzoomed coordinates. Returns horizontal offset relative to the page
 * content area (column origin) and vertical offset relative to the
 * containing paragraph — matches the OOXML attrs the resolver writes
 * (`relativeFrom: 'column'` / `relativeFrom: 'paragraph'`).
 *
 * `zoom` defaults to 1; pass the editor's current zoom factor when the
 * pages container has a CSS scale applied so `getBoundingClientRect` deltas
 * are converted back to authored-pixel space before going to EMU.
 *
 * Returns undefined for non-inline images or detached DOM.
 */
export function captureInlinePositionEmu(
  imageEl: HTMLElement,
  zoom = 1,
): { horizontalEmu: number; verticalEmu: number } | undefined {
  if (!imageEl.classList.contains(LAYOUT_IMAGE_CLASSES.runImage)) {
    return undefined;
  }
  const pageContent = imageEl.closest<HTMLElement>(`.${LAYOUT_IMAGE_CLASSES.pageContent}`);
  const paragraph = imageEl.closest<HTMLElement>(`.${LAYOUT_IMAGE_CLASSES.paragraph}`);
  if (!pageContent || !paragraph) {
    return undefined;
  }
  const imgRect = imageEl.getBoundingClientRect();
  const pageRect = pageContent.getBoundingClientRect();
  const paraRect = paragraph.getBoundingClientRect();
  const safeZoom = zoom > 0 ? zoom : 1;
  return {
    horizontalEmu: Math.round(pixelsToEmu((imgRect.left - pageRect.left) / safeZoom)),
    verticalEmu: Math.round(pixelsToEmu((imgRect.top - paraRect.top) / safeZoom)),
  };
}

// ============================================================================
// User-facing layout choice ↔ PM image attrs
// ============================================================================

/**
 * Map an image's current OOXML attrs onto the menu's choice vocabulary so the
 * menu can highlight the active option. Returns null for `topAndBottom` and
 * any unknown wrap type — those don't have a directional menu entry.
 *
 * `cssFloat` accepts both `null` and `undefined` so framework adapters can
 * use either as their "unset" sentinel without an extra normalisation step.
 */
export function deriveLayoutChoice(
  wrapType: WrapType,
  cssFloat?: ImageAttrs["cssFloat"] | null,
): ImageLayoutTarget | null {
  if (wrapType === "inline") {
    return "inline";
  }
  if (wrapType === "behind") {
    return "behind";
  }
  if (wrapType === "inFront") {
    return "inFront";
  }
  if (wrapType === "square" || wrapType === "tight" || wrapType === "through") {
    return cssFloat === "right" ? "squareRight" : "squareLeft";
  }
  return null;
}

// ============================================================================
// Layout option schema (config without per-framework presentation)
// ============================================================================

/**
 * Hint to the framework's icon registry for which Material Symbol — or
 * equivalent — to render alongside each option. Bindings own the icon
 * component itself.
 */
export type ImageLayoutIconHint = "inline" | "squareLeft" | "squareRight" | "behind" | "inFront";

export type ImageLayoutOptionDef = {
  /** Choice value — what gets dispatched on click. */
  choice: ImageLayoutTarget;
  /** i18n key under `imageWrap.menu.*`. */
  i18nLabelKey: string;
  /** i18n key under `imageWrap.menuDesc.*`. */
  i18nDescKey: string;
  /** Hint for the framework's icon registry. */
  iconHint: ImageLayoutIconHint;
};

/** Mirrors Word's Wrap Text menu — five directional options. */
export const IMAGE_LAYOUT_OPTIONS: readonly ImageLayoutOptionDef[] = [
  {
    choice: "inline",
    i18nLabelKey: "inLineWithText",
    i18nDescKey: "inLineWithText",
    iconHint: "inline",
  },
  {
    choice: "squareLeft",
    i18nLabelKey: "squareLeft",
    i18nDescKey: "squareLeft",
    iconHint: "squareLeft",
  },
  {
    choice: "squareRight",
    i18nLabelKey: "squareRight",
    i18nDescKey: "squareRight",
    iconHint: "squareRight",
  },
  {
    choice: "behind",
    i18nLabelKey: "behindText",
    i18nDescKey: "behindText",
    iconHint: "behind",
  },
  {
    choice: "inFront",
    i18nLabelKey: "inFrontOfText",
    i18nDescKey: "inFrontOfText",
    iconHint: "inFront",
  },
] as const;

/**
 * Whether a given option is enabled for an image with the given current wrap
 * type. Every option stays clickable — picking the option that matches the
 * current state is a no-op (the PM command early-returns), which matches
 * Word's behavior. Kept around for forward-compatibility (e.g. future
 * read-only mode), but currently always returns true.
 */
export function isImageLayoutOptionEnabled(
  _option: ImageLayoutOptionDef,
  _currentWrapType: WrapType,
): boolean {
  return true;
}

// ============================================================================
// Toolbar-vocabulary mapping
// ============================================================================

/**
 * Translate the legacy toolbar wrap-type vocabulary (`wrapLeft` / `wrapRight`
 * / `square` / `tight` / `through` / `topAndBottom` / `behind` / `inFront` /
 * `inline`) into an `ImageLayoutTarget` so toolbar dispatch shares the same PM
 * command path as the right-click menu.
 *
 * Returns `undefined` for unknown values; callers should treat that as
 * "no-op".
 */
export function toolbarValueToLayoutTarget(value: string): ImageLayoutTarget | undefined {
  switch (value) {
    case "inline":
      return "inline";
    case "square":
    case "tight":
    case "through":
      return "squareLeft";
    case "topAndBottom":
      return "topAndBottom";
    case "behind":
      return "behind";
    case "inFront":
      return "inFront";
    case "wrapRight":
      return "squareLeft";
    case "wrapLeft":
      return "squareRight";
    default:
      return undefined;
  }
}

// ============================================================================
// Block lookup
// ============================================================================

/**
 * Build the painter's `block.id → { block, measure }` lookup from the parallel
 * blocks/measures arrays. Shared by both adapters' paint step.
 */
export function buildBlockLookup(blocks: FlowBlock[], measures: Measure[]): BlockLookup {
  const lookup: BlockLookup = new Map();
  for (let i = 0; i < blocks.length; i++) {
    const block = blocks[i];
    const measure = measures[i];
    if (block && measure) {
      lookup.set(String(block.id), { block, measure });
    }
  }
  return lookup;
}
