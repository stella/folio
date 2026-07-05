/**
 * Stamp `data-sdt-*` attributes on a painted block element from the
 * fragment's `sdtGroups` projection.
 *
 * The innermost group drives the boundary chrome and the widget delegation
 * layer; the full outer→inner stack is exposed as a JSON list so a hover
 * surface or addressing API can walk it. The painter side is intentionally
 * passive — CSS draws the boundary based on `data-sdt-boundary`, and the
 * widget layer reads the rest to render interactive triggers.
 */

import type { Node as PMNode } from "prosemirror-model";

import type { SdtGroup } from "../layout-engine/types";

export function applySdtDataAttrs(
  el: HTMLElement,
  sdtGroups: readonly SdtGroup[] | undefined,
): void {
  if (!sdtGroups || sdtGroups.length === 0) {
    return;
  }
  el.dataset["sdtBoundary"] = "true";

  const innermost = sdtGroups.at(-1);
  if (innermost) {
    el.dataset["sdtId"] = innermost.id;
    el.dataset["sdtPmPos"] = String(innermost.pmPos);
    el.dataset["sdtType"] = innermost.sdtType;
    if (innermost.position) {
      el.dataset["sdtPosition"] = innermost.position;
    }
    if (innermost.alias) {
      el.dataset["sdtAlias"] = innermost.alias;
    }
    if (innermost.tag) {
      el.dataset["sdtTag"] = innermost.tag;
    }
    if (innermost.sdtId !== undefined) {
      el.dataset["sdtOoxmlId"] = String(innermost.sdtId);
    }
    if (innermost.lock) {
      el.dataset["sdtLock"] = innermost.lock;
    }
    if (innermost.showingPlaceholder) {
      el.dataset["sdtShowingPlaceholder"] = "true";
    }
    if (innermost.checked !== undefined) {
      el.dataset["sdtChecked"] = String(innermost.checked);
    }
    if (innermost.dateFormat) {
      el.dataset["sdtDateFormat"] = innermost.dateFormat;
    }
    if (innermost.listItemsJson) {
      el.dataset["sdtListItems"] = innermost.listItemsJson;
    }
  }

  // Outer→inner stack for callers that need to walk it (addressing API,
  // widget delegation layer).
  el.dataset["sdtStack"] = JSON.stringify(
    sdtGroups.map((g) => ({
      id: g.id,
      sdtType: g.sdtType,
      tag: g.tag ?? null,
      alias: g.alias ?? null,
    })),
  );
}

/**
 * Identities of every block-level content control (`w:sdt`) that encloses the
 * given selection — used to keep a control's boundary visible while the caret
 * is inside it (Word-style focus), independent of mouse hover. Both ends of a
 * range are collected so a selection straddling a control still lights it up.
 *
 * Ported from eigenpal/docx-editor `layout-painter/sdtBoundary.ts`. Upstream
 * keyed the set on `sdt@${pos}` to match per-control boundary boxes; this fork
 * stamps the SDT node's PM position as `data-sdt-pm-pos` on the painted block
 * elements (see `applySdtDataAttrs`), so the identity here is the stringified
 * `$pos.before(d)` — which is exactly that PM position — keeping the returned
 * ids in lockstep with what `applySdtFocus` hit-tests against.
 */
export function enclosingSdtGroupIds(doc: PMNode, from: number, to: number): Set<string> {
  const ids = new Set<string>();
  const max = doc.content.size;
  const collect = (pos: number): void => {
    const $pos = doc.resolve(Math.max(0, Math.min(pos, max)));
    for (let d = 1; d <= $pos.depth; d++) {
      if ($pos.node(d).type.name === "blockSdt") {
        ids.add(String($pos.before(d)));
      }
    }
  };
  collect(from);
  if (to !== from) {
    collect(to);
  }
  return ids;
}

/**
 * Toggle the `.is-focused` reveal class on every painted block whose enclosing
 * content control encloses the caret. Kept separate from any hover-driven
 * reveal so the two paths never clear each other. Matches the fork's
 * `data-sdt-pm-pos` stamping (upstream matched `.layout-block-sdt-box`'s
 * `data-sdt-group-id`, which this fork does not paint).
 */
export function applySdtFocus(container: HTMLElement, focusedIds: Set<string>): void {
  const boxes = container.querySelectorAll<HTMLElement>("[data-sdt-boundary]");
  for (const box of boxes) {
    const id = box.dataset["sdtPmPos"];
    const on = id != null && focusedIds.has(id);
    if (on !== box.classList.contains("is-focused")) {
      box.classList.toggle("is-focused", on);
    }
  }
}
