/**
 * Anonymization Match Plugin
 *
 * Owns the "what to highlight" half of the anonymization
 * overlay: the host pushes a list of terms via the
 * `setAnonymizationTermsMeta` meta, this plugin scans text
 * nodes for each surface form, and exposes the resulting
 * match ranges through its state.
 *
 * The plugin does **not** produce ProseMirror decorations.
 * Folio's editor lives off-screen (HiddenProseMirror) and PM
 * decorations never reach the visible paged DOM. The visible
 * highlights are painted by {@link AnonymizationRectsOverlay},
 * which reads the match ranges from this plugin's state and
 * projects them through `selectionToRects` to coordinates on
 * the paged canvas. Single source of truth for the ranges,
 * one painter.
 */

import type { Node as PMNode } from "prosemirror-model";
import { Plugin } from "prosemirror-state";
import type { EditorState, PluginKey, PluginSpec } from "prosemirror-state";

import { createPluginKeyOwner, type PluginKeyOwner } from "../pluginKeyOwner";
import { collectBlockChunks, joinChunks, offsetToDocPos } from "./pmTextScan";

export type AnonymizationTerm = {
  /** Canonical surface form, displayed in tooltips. */
  canonical: string;
  /** Label slug (e.g. "person", "organization"). */
  label: string;
  /** Optional alternate surface forms also matched verbatim. */
  variants?: readonly string[];
};

export type AnonymizationMatch = {
  /** Inclusive PM doc position of the match start. */
  from: number;
  /** Exclusive PM doc position of the match end. */
  to: number;
  label: string;
  canonical: string;
};

const SET_META = "set";

type AnonymizationDecorationState = {
  terms: readonly AnonymizationTerm[];
  matches: readonly AnonymizationMatch[];
};

export const anonymizationDecorationsKey: Readonly<{
  key: string;
  createKey: () => PluginKey<AnonymizationDecorationState>;
  owns: (plugin: Plugin<AnonymizationDecorationState>) => boolean;
  get: (state: EditorState) => Plugin<AnonymizationDecorationState> | undefined;
  getState: (state: EditorState) => AnonymizationDecorationState | undefined;
}> = createPluginKeyOwner<AnonymizationDecorationState>(
  "anonymizationDecorations",
);

export const slugAnonymizationLabel = (label: string): string =>
  label.toLowerCase().replace(/[^a-z0-9]+/gu, "-");

type CompiledTerm = {
  regex: RegExp;
  term: AnonymizationTerm;
};

const buildMatcher = (terms: readonly AnonymizationTerm[]): CompiledTerm[] => {
  const escapeChar = (value: string): string =>
    value
      .replaceAll(/[\\^$.*+?()[\]{}|]/gu, "\\$&")
      // A typed whitespace run matches the non-breaking, narrow
      // no-break and figure spaces typography commonly inserts —
      // purely a normalisation concern, not language logic.
      .replaceAll(/\s+/gu, "\\s+");

  // One regex per term (canonical + variants). The plugin only
  // does literal, word-bounded matching; morphology and other
  // form expansion belong upstream in the anonymisation package
  // (variants on the gazetteer entry).
  const compiled: CompiledTerm[] = [];
  const seen = new Set<string>();
  for (const term of terms) {
    const surfaces = [term.canonical, ...(term.variants ?? [])].filter(
      (surface) => surface.length > 0,
    );
    if (surfaces.length === 0) {
      continue;
    }
    surfaces.sort((a, b) => b.length - a.length);
    const key = surfaces.join("|").toLowerCase();
    if (seen.has(key)) {
      continue;
    }
    seen.add(key);
    const alternation = surfaces.map(escapeChar).join("|");
    compiled.push({
      term,
      regex: new RegExp(
        `(?<![\\p{L}\\p{N}])(?:${alternation})(?![\\p{L}\\p{N}])`,
        "giu",
      ),
    });
  }
  return compiled;
};

const buildMatches = (
  doc: PMNode,
  terms: readonly AnonymizationTerm[],
): AnonymizationMatch[] => {
  const compiled = buildMatcher(terms);
  if (compiled.length === 0) {
    return [];
  }
  const matches: AnonymizationMatch[] = [];
  for (const chunks of collectBlockChunks(doc)) {
    const joined = joinChunks(chunks);
    for (const { regex, term } of compiled) {
      regex.lastIndex = 0;
      let match: RegExpExecArray | null = regex.exec(joined);
      while (match !== null) {
        const from = offsetToDocPos(chunks, match.index);
        const to = offsetToDocPos(chunks, match.index + match[0].length, "end");
        matches.push({
          from,
          to,
          label: term.label,
          canonical: term.canonical,
        });
        if (regex.lastIndex === match.index) {
          regex.lastIndex += 1;
        }
        match = regex.exec(joined);
      }
    }
  }
  // De-duplicate overlapping matches: shorter spans nested inside
  // longer ones get dropped so a single occurrence of "First Last"
  // does not paint twice (once for the full name, once for the
  // canonical word inside it).
  matches.sort((a, b) => (a.from !== b.from ? a.from - b.from : b.to - a.to));
  const result: AnonymizationMatch[] = [];
  let cursor = -1;
  for (const m of matches) {
    if (m.from < cursor) {
      continue;
    }
    result.push(m);
    cursor = m.to;
  }
  return result;
};

/**
 * Plugin that keeps a list of anonymization-term match ranges
 * synced with the document. Host pushes terms via
 * {@link setAnonymizationTermsMeta}; the paged-editor reads
 * the matches and paints the overlay.
 */
export type AnonymizationPluginOptions = {
  /**
   * Push-side bridge for hosts that need to mirror the current
   * match list outside the PM state tree (e.g. an inspector
   * facet showing a "N highlighted" counter, or per-doc analytics).
   * Called on plugin init, after every transaction that changes
   * the match set, and on plugin teardown (with an empty list).
   * The same list is reachable via `getAnonymizationMatches`;
   * use the callback when you need updates without polling.
   */
  onMatchesChange?: (matches: readonly AnonymizationMatch[]) => void;
};

export const createAnonymizationDecorationsPlugin = ({
  onMatchesChange,
}: AnonymizationPluginOptions = {}): Plugin<AnonymizationDecorationState> => {
  const spec: PluginSpec<AnonymizationDecorationState> = {
    key: anonymizationDecorationsKey.createKey(),
    state: {
      init(): AnonymizationDecorationState {
        return { terms: [], matches: [] };
      },
      apply(tr, prev, _oldState, newState): AnonymizationDecorationState {
        const setMeta = tr.getMeta(anonymizationDecorationsKey) as
          | { type: typeof SET_META; terms: readonly AnonymizationTerm[] }
          | undefined;

        if (setMeta?.type === SET_META) {
          return {
            terms: setMeta.terms,
            matches: buildMatches(newState.doc, setMeta.terms),
          };
        }

        if (tr.docChanged) {
          // Doc edits move text around; rebuild ranges from scratch
          // rather than try to map regex matches through a mapping.
          return {
            terms: prev.terms,
            matches: buildMatches(newState.doc, prev.terms),
          };
        }
        return prev;
      },
    },
  };
  if (onMatchesChange) {
    spec.view = (view) => {
      // Emit the initial match list (init() ran above with the
      // empty doc-and-terms baseline; for restored sessions
      // matches may already be populated).
      let last = anonymizationDecorationsKey.getState(view.state)?.matches;
      if (last) {
        onMatchesChange(last);
      }
      return {
        update(updatedView) {
          const next = anonymizationDecorationsKey.getState(
            updatedView.state,
          )?.matches;
          if (next && next !== last) {
            last = next;
            onMatchesChange(next);
          }
        },
        destroy() {
          onMatchesChange([]);
        },
      };
    };
  }
  return new Plugin<AnonymizationDecorationState>(spec);
};

export const setAnonymizationTermsMeta = (
  terms: readonly AnonymizationTerm[],
): {
  key: PluginKeyOwner<AnonymizationDecorationState>;
  payload: { type: typeof SET_META; terms: readonly AnonymizationTerm[] };
} => ({
  key: anonymizationDecorationsKey,
  payload: { type: SET_META, terms },
});

/** Read anonymization matches owned by this editor state. */
export const getAnonymizationMatches = (
  state: EditorState,
): readonly AnonymizationMatch[] =>
  anonymizationDecorationsKey.getState(state)?.matches ?? [];
