/**
 * Apply the visual cell-selection highlight class onto rendered table cells.
 *
 * The layout painter renders pages as static DOM with `data-pm-start` on each
 * table cell. This walk translates a ProseMirror `CellSelection` into the
 * `.layout-table-cell-selected` class on the matching painted cells. Shared by
 * the React and Vue adapters.
 */

import type { EditorState } from "prosemirror-state";

type CellSelectionLike = {
  $anchorCell?: unknown;
  forEachCell?: (cb: (node: { nodeSize: number }, pos: number) => void) => void;
};

/**
 * Apply the `.layout-table-cell-selected` class to painted layout cells
 * matching a CellSelection in the PM state. Clears the class everywhere
 * (within scope) first so toggling off (or moving to a TextSelection) erases
 * prior highlights.
 *
 * Duck-types CellSelection via `$anchorCell` / `forEachCell` rather than
 * `instanceof` to dodge bundling issues across `prosemirror-tables` copies
 * (the same trick used inline in the selection-overlay update).
 */
export function applyCellSelectionHighlight(
  pagesContainer: HTMLElement,
  state: EditorState,
  options: { scope?: "body" | "header" | "footer" } = {},
): void {
  const scope = options.scope ?? "body";
  // The selector that limits which cells this call can highlight. Header and
  // footer cells live in `.layout-page-header` / `.layout-page-footer`
  // (separate PM docs), body cells in `.layout-page-content`. PM positions
  // overlap across all three docs, so we MUST scope the walk to the matching
  // tree — a footer CellSelection at pos 100 would otherwise also light up a
  // header cell at `data-pm-start="100"` (#671), and a body selection would
  // bleed into both.
  const scopeClass = scope === "body" ? "layout-page-content" : `layout-page-${scope}`;
  const scopeSelector = `.${scopeClass} .layout-table-cell`;

  // Only clear highlights inside this scope so the body call doesn't wipe
  // HF highlights (and vice versa), and the header call doesn't wipe the
  // footer's.
  const prevSelected = pagesContainer.querySelectorAll(
    `.${scopeClass} .layout-table-cell-selected`,
  );
  for (const el of Array.from(prevSelected)) {
    el.classList.remove("layout-table-cell-selected");
  }

  // SAFETY: prosemirror-tables' `CellSelection` cannot be identified with
  // `instanceof` across bundled copies, so duck-type its structural shape.
  const sel = state.selection as unknown as CellSelectionLike;
  const isCellSel = "$anchorCell" in sel && typeof sel.forEachCell === "function";
  if (!isCellSel || !sel.forEachCell) return;

  const selectedRanges: Array<[number, number]> = [];
  sel.forEachCell((node, pos) => {
    selectedRanges.push([pos, pos + node.nodeSize]);
  });

  const allCells = pagesContainer.querySelectorAll<HTMLElement>(scopeSelector);
  for (const cellEl of Array.from(allCells)) {
    const pmStartAttr = cellEl.dataset["pmStart"];
    if (pmStartAttr === undefined) continue;
    const pmPos = Number(pmStartAttr);
    for (const [start, end] of selectedRanges) {
      if (pmPos >= start && pmPos < end) {
        cellEl.classList.add("layout-table-cell-selected");
        break;
      }
    }
  }
}
