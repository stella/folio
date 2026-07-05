/**
 * Selection highlight utilities.
 *
 * Custom visual highlighting for text selection that spans multiple runs, where
 * the browser's native `::selection` pseudo-element renders inconsistently
 * across elements with different backgrounds. Framework-agnostic DOM helpers:
 * the adapter layer wires them to reactive state.
 * @packageDocumentation
 * @public
 */

/** Highlight rectangle representing a selected region, in pixels. */
export type HighlightRect = {
  left: number;
  top: number;
  width: number;
  height: number;
}

/** Selection highlight configuration. */
export type SelectionHighlightConfig = {
  /** Background color for selection. */
  backgroundColor: string;
  /** Optional border color for selection. */
  borderColor?: string;
  /** Optional border radius, in pixels. */
  borderRadius?: number;
  /** Z-index for overlay. */
  zIndex?: number;
  /** Opacity for highlight. */
  opacity?: number;
  /** CSS `mix-blend-mode` for the overlay. */
  mixBlendMode?: string;
}

/** Default selection highlight style (matches Word / Google Docs). */
export const DEFAULT_SELECTION_STYLE: SelectionHighlightConfig = {
  backgroundColor: "rgba(26, 115, 232, 0.3)",
  borderRadius: 0,
  zIndex: 0,
  opacity: 1,
  mixBlendMode: "multiply",
};

/**
 * Get all selection rectangles from the current DOM selection.
 *
 * Uses `getClientRects()` for accurate rectangles even when the selection spans
 * multiple inline elements. When a container is supplied, rectangles are
 * returned relative to it (and an out-of-container selection yields none).
 */
function getSelectionRects(containerElement?: HTMLElement | null): HighlightRect[] {
  const selection = window.getSelection();
  if (!selection || selection.rangeCount === 0 || selection.isCollapsed) {
    return [];
  }

  const range = selection.getRangeAt(0);

  if (containerElement && !containerElement.contains(range.commonAncestorContainer)) {
    return [];
  }

  let offsetLeft = 0;
  let offsetTop = 0;
  if (containerElement) {
    const containerRect = containerElement.getBoundingClientRect();
    offsetLeft = containerRect.left + containerElement.scrollLeft;
    offsetTop = containerRect.top + containerElement.scrollTop;
  }

  const rects: HighlightRect[] = [];
  for (const rect of range.getClientRects()) {
    // Skip zero-size rects (these can occur at line breaks).
    if (rect.width === 0 && rect.height === 0) {
      continue;
    }
    rects.push({
      left: rect.left - offsetLeft,
      top: rect.top - offsetTop,
      width: rect.width,
      height: rect.height,
    });
  }
  return rects;
}

/**
 * Merge adjacent or overlapping rectangles to reduce the number of highlight
 * elements and produce a cleaner visual appearance.
 */
function mergeAdjacentRects(rects: HighlightRect[], tolerance = 2): HighlightRect[] {
  if (rects.length <= 1) {
    return rects;
  }

  const sorted = [...rects].sort((a, b) =>
    Math.abs(a.top - b.top) < tolerance ? a.left - b.left : a.top - b.top,
  );

  const first = sorted[0];
  if (!first) {
    return rects;
  }

  const merged: HighlightRect[] = [];
  let current: HighlightRect = { ...first };

  for (let i = 1; i < sorted.length; i++) {
    const rect = sorted[i];
    if (!rect) {
      continue;
    }

    const sameLine = Math.abs(rect.top - current.top) < tolerance;
    const adjacent = rect.left <= current.left + current.width + tolerance;

    if (sameLine && adjacent) {
      const newRight = Math.max(current.left + current.width, rect.left + rect.width);
      current.width = newRight - current.left;
      current.height = Math.max(current.height, rect.height);
    } else {
      merged.push(current);
      current = { ...rect };
    }
  }

  merged.push(current);
  return merged;
}

/** Get selection rectangles with adjacent-rect merging applied. */
export function getMergedSelectionRects(containerElement?: HTMLElement | null): HighlightRect[] {
  return mergeAdjacentRects(getSelectionRects(containerElement));
}

/** Whether there is an active, non-collapsed text selection. */
export function hasActiveSelection(): boolean {
  const selection = window.getSelection();
  return selection !== null && !selection.isCollapsed && selection.rangeCount > 0;
}

/** Whether the current selection is within a specific element. */
export function isSelectionWithin(element: HTMLElement): boolean {
  const selection = window.getSelection();
  if (!selection || selection.rangeCount === 0) {
    return false;
  }
  return element.contains(selection.getRangeAt(0).commonAncestorContainer);
}

const SELECTION_STYLE_ELEMENT_ID = "docx-selection-styles";
let injectedStyleElement: HTMLStyleElement | null = null;

/** Remove any previously injected selection styles. */
function removeSelectionStyles(): void {
  if (injectedStyleElement) {
    injectedStyleElement.remove();
    injectedStyleElement = null;
  }
  // Also remove by id in case the reference was lost (e.g. HMR).
  document.getElementById(SELECTION_STYLE_ELEMENT_ID)?.remove();
}

/** Inject selection highlight CSS into the document head. */
export function injectSelectionStyles(config: SelectionHighlightConfig = DEFAULT_SELECTION_STYLE): void {
  removeSelectionStyles();

  const css = `
    /* DOCX Editor Selection Highlighting */

    .docx-editor [contenteditable="true"]::selection,
    .docx-editor [contenteditable="true"] *::selection,
    .docx-run-editable::selection,
    .docx-run-editable *::selection {
      background-color: ${config.backgroundColor} !important;
      color: inherit !important;
    }

    .docx-editor [contenteditable="true"]::-moz-selection,
    .docx-editor [contenteditable="true"] *::-moz-selection,
    .docx-run-editable::-moz-selection,
    .docx-run-editable *::-moz-selection {
      background-color: ${config.backgroundColor} !important;
      color: inherit !important;
    }

    .docx-run-highlighted::selection,
    .docx-run-highlighted *::selection {
      background-color: rgba(26, 115, 232, 0.5) !important;
    }

    .docx-run-highlighted::-moz-selection,
    .docx-run-highlighted *::-moz-selection {
      background-color: rgba(26, 115, 232, 0.5) !important;
    }

    .docx-run-dark-bg::selection,
    .docx-run-dark-bg *::selection {
      background-color: rgba(100, 181, 246, 0.5) !important;
    }

    .docx-run-dark-bg::-moz-selection,
    .docx-run-dark-bg *::-moz-selection {
      background-color: rgba(100, 181, 246, 0.5) !important;
    }

    .docx-selection-highlight {
      background-color: ${config.backgroundColor};
      ${config.borderRadius ? `border-radius: ${config.borderRadius}px;` : ""}
      ${config.mixBlendMode ? `mix-blend-mode: ${config.mixBlendMode};` : ""}
    }

    .docx-find-highlight {
      background-color: rgba(255, 235, 59, 0.5);
      border-radius: 2px;
    }

    .docx-find-highlight-current {
      background-color: rgba(255, 152, 0, 0.6);
      border-radius: 2px;
      outline: 2px solid rgba(255, 152, 0, 0.8);
    }

    .docx-ai-selection-preview {
      background-color: rgba(156, 39, 176, 0.2);
      border-bottom: 2px dashed rgba(156, 39, 176, 0.6);
    }
  `;

  injectedStyleElement = document.createElement("style");
  injectedStyleElement.id = SELECTION_STYLE_ELEMENT_ID;
  injectedStyleElement.textContent = css;
  document.head.appendChild(injectedStyleElement);
}

/** Whether selection styles are currently injected. */
export function areSelectionStylesInjected(): boolean {
  return injectedStyleElement !== null || document.getElementById(SELECTION_STYLE_ELEMENT_ID) !== null;
}
