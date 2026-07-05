/**
 * ProseMirror-position → page-index helpers for the layout engine.
 *
 * Framework-neutral utilities shared by every rendering adapter: locating the
 * page whose layout fragments cover a PM position, plus the exhaustiveness
 * guard for `FlowBlock`-shaped switches.
 */

import type { Layout } from "./types";

/**
 * Exhaustiveness guard for `FlowBlock`-shaped switches. Call from the
 * `default` arm with the still-typed value; TypeScript will refuse to
 * compile if any variant of `FlowBlock` was missed. The thrown error
 * names the calling site so runtime failures (e.g. an old adapter
 * compiled against a newer core) point future debuggers at the contract.
 */
export function assertExhaustiveFlowBlock(block: never, site: string): never {
  const kind = (block as { kind?: string }).kind ?? "<unknown>";
  throw new Error(
    `${site}: unhandled FlowBlock kind "${kind}". ` +
      "Add the case alongside the other FlowBlock switches (see types.ts).",
  );
}

/**
 * Page index (0-based) whose layout fragments cover `pmPos`, or null if none.
 * Used when the painted DOM may not yet have `[data-pm-start]` for this position (virtualization).
 *
 * Range semantics: `[pmStart, pmEnd)` — half-open, matching ProseMirror's
 * `pos + nodeSize` convention. Boundary positions belong to the next fragment,
 * so when a fragment ends at the same position the next one starts, the next
 * fragment wins (avoids returning the previous page for the start of the
 * next paragraph).
 */
export function findPageIndexContainingPmPos(layout: Layout, pmPos: number): number | null {
  for (let pi = 0; pi < layout.pages.length; pi++) {
    const page = layout.pages[pi];
    if (!page) continue;
    for (const frag of page.fragments) {
      if (frag.pmStart == null) continue;
      const start = frag.pmStart;
      // Default span of 1 only when pmEnd is missing — matches a caret-only
      // position (cursor between two atoms). Fragments with explicit pmEnd
      // use it as the exclusive upper bound.
      const end = frag.pmEnd ?? start + 1;
      if (pmPos >= start && pmPos < end) {
        return pi;
      }
    }
  }
  return null;
}
