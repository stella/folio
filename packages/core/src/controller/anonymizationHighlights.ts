/**
 * Anonymization-term highlights: where a host's terms occur in the text the
 * editor shows, in every story it paints (body with its table cells,
 * headers, footers, footnotes, endnotes).
 *
 * Highlights are view state. They never touch the document, its history or
 * its saved state; they are derived from the shown text whenever it is read.
 * Each paragraph's matches are kept against the paragraph's own editor node,
 * so a read after an edit rescans only the paragraphs the edit replaced (an
 * editor document shares every node an edit did not reach), whoever made the
 * edit: the person, an undo, or a peer in a shared room. A change of terms
 * rescans everything.
 *
 * Matching is literal, case-insensitive and word-bounded over a paragraph's
 * joined text, so a term that crosses formatting runs, a comment anchor or a
 * tracked insertion's edge still matches. Text inside a tracked deletion is
 * shown text too and is matched like any other.
 */

import type { Node as PMNode } from "prosemirror-model";

import {
  OP_STORIES,
  type OpStory,
  type TextPosition,
} from "@stll/docx-core/ops";

import {
  collectTextblockChunks,
  joinChunks,
  offsetToDocPos,
} from "../prosemirror/plugins/pmTextScan";
import type {
  FolioAnonymizationMatch,
  FolioAnonymizationTerm,
} from "../types/anonymization";

export type {
  FolioAnonymizationMatch,
  FolioAnonymizationTerm,
} from "../types/anonymization";

/** One occurrence as an editor shows it: positions in its story's editor document. */
export type AnonymizationShownMatch = {
  story: OpStory;
  from: number;
  to: number;
  label: string;
  canonical: string;
};

/** A story's editor document and how its positions read as text positions. */
export type AnonymizationStorySource = {
  story: OpStory;
  doc: PMNode;
  addressAt: (position: number) => TextPosition | null;
};

export type AnonymizationHighlightsView = {
  /** Every occurrence, in story order, for painting. */
  shown: readonly AnonymizationShownMatch[];
  /**
   * Every occurrence as text positions. The same array until an occurrence
   * appears, disappears or moves within its paragraph.
   */
  matches: readonly FolioAnonymizationMatch[];
};

export type AnonymizationHighlights = {
  /** Read the stories `sources` gives only when there is a term to find. */
  read: (
    terms: readonly FolioAnonymizationTerm[],
    sources: () => readonly AnonymizationStorySource[],
  ) => AnonymizationHighlightsView;
};

/** No terms: highlights nothing. One shared value, so reads keep their scans. */
export const NO_ANONYMIZATION_TERMS: readonly FolioAnonymizationTerm[] =
  Object.freeze([]);

export const slugAnonymizationLabel = (label: string): string =>
  label.toLowerCase().replace(/[^a-z0-9]+/gu, "-");

type CompiledTerm = { regex: RegExp; term: FolioAnonymizationTerm };

const compileTerms = (
  terms: readonly FolioAnonymizationTerm[],
): CompiledTerm[] => {
  const escape = (value: string): string =>
    value
      .replaceAll(/[\\^$.*+?()[\]{}|]/gu, "\\$&")
      // A typed whitespace run matches the non-breaking, narrow no-break and
      // figure spaces typography inserts.
      .replaceAll(/\s+/gu, "\\s+");
  const compiled: CompiledTerm[] = [];
  const seen = new Set<string>();
  for (const term of terms) {
    const surfaces = [term.canonical, ...(term.variants ?? [])].filter(
      (surface) => surface.length > 0,
    );
    if (surfaces.length === 0) continue;
    // Longer surfaces first, so an alternation prefers the longest form.
    surfaces.sort((a, b) => b.length - a.length);
    const key = surfaces.join("|").toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    compiled.push({
      term,
      regex: new RegExp(
        `(?<![\\p{L}\\p{N}])(?:${surfaces.map(escape).join("|")})(?![\\p{L}\\p{N}])`,
        "giu",
      ),
    });
  }
  return compiled;
};

/** A match inside one paragraph, positions relative to the paragraph node's start. */
type BlockMatch = {
  from: number;
  to: number;
  term: FolioAnonymizationTerm;
  address: FolioAnonymizationMatch | null;
};

const scanBlock = (
  node: PMNode,
  position: number,
  compiled: readonly CompiledTerm[],
  addressAt: AnonymizationStorySource["addressAt"],
): readonly BlockMatch[] => {
  const found: Omit<BlockMatch, "address">[] = [];
  // Relative to the paragraph: chunk positions as if the node started at 0.
  for (const chunks of collectTextblockChunks(node, 0)) {
    const joined = joinChunks(chunks);
    for (const { regex, term } of compiled) {
      regex.lastIndex = 0;
      for (
        let match = regex.exec(joined);
        match !== null;
        match = regex.exec(joined)
      ) {
        found.push({
          from: offsetToDocPos(chunks, match.index),
          to: offsetToDocPos(chunks, match.index + match[0].length, "end"),
          term,
        });
        if (regex.lastIndex === match.index) regex.lastIndex += 1;
      }
    }
  }
  if (found.length === 0) return EMPTY_BLOCK;
  // A shorter occurrence nested in a longer one paints once, as the longer.
  found.sort((a, b) => (a.from !== b.from ? a.from - b.from : b.to - a.to));
  const kept: BlockMatch[] = [];
  let cursor = -1;
  for (const match of found) {
    if (match.from < cursor) continue;
    cursor = match.to;
    const start = addressAt(position + match.from);
    const end = addressAt(position + match.to);
    kept.push({
      ...match,
      address:
        start === null || end === null
          ? null
          : {
              start,
              end,
              label: match.term.label,
              canonical: match.term.canonical,
            },
    });
  }
  return kept;
};

const EMPTY_BLOCK: readonly BlockMatch[] = Object.freeze([]);
/** Nothing highlighted. */
export const NO_ANONYMIZATION_HIGHLIGHTS: AnonymizationHighlightsView =
  Object.freeze({
    shown: Object.freeze([]),
    matches: Object.freeze([]),
  });

const storyKey = (story: OpStory): string => {
  if (story === OP_STORIES.MAIN) return story;
  if ("rId" in story) return `${story.kind}:${story.rId}`;
  return `${story.kind}:${String(story.id)}`;
};

/**
 * Collect one story's occurrences, rescanning only the paragraphs `scanned`
 * does not hold yet.
 */
const readStory = ({
  source,
  scanned,
  compiled,
  blocks,
  shown,
}: {
  source: AnonymizationStorySource;
  scanned: WeakMap<PMNode, readonly BlockMatch[]>;
  compiled: readonly CompiledTerm[];
  blocks: (readonly BlockMatch[])[];
  shown: AnonymizationShownMatch[];
}): void => {
  source.doc.descendants((node, position) => {
    if (!node.isTextblock) return true;
    let found = scanned.get(node);
    if (found === undefined) {
      found = scanBlock(node, position, compiled, source.addressAt);
      scanned.set(node, found);
    }
    if (found.length > 0) {
      blocks.push(found);
      for (const match of found)
        shown.push({
          story: source.story,
          from: position + match.from,
          to: position + match.to,
          label: match.term.label,
          canonical: match.term.canonical,
        });
    }
    return false;
  });
};

/** The highlight owner of one editor. */
export const createAnonymizationHighlights = (): AnonymizationHighlights => {
  let terms: readonly FolioAnonymizationTerm[] = [];
  let compiled: readonly CompiledTerm[] = [];
  let scans = new Map<string, WeakMap<PMNode, readonly BlockMatch[]>>();
  let last: {
    blocks: readonly (readonly BlockMatch[])[];
    matches: readonly FolioAnonymizationMatch[];
  } = { blocks: [], matches: NO_ANONYMIZATION_HIGHLIGHTS.matches };

  return {
    read: (nextTerms, sources) => {
      if (nextTerms !== terms) {
        terms = nextTerms;
        compiled = compileTerms(nextTerms);
        scans = new Map();
      }
      if (compiled.length === 0) {
        last = { blocks: [], matches: NO_ANONYMIZATION_HIGHLIGHTS.matches };
        return NO_ANONYMIZATION_HIGHLIGHTS;
      }
      const blocks: (readonly BlockMatch[])[] = [];
      const shown: AnonymizationShownMatch[] = [];
      for (const source of sources()) {
        const key = storyKey(source.story);
        let scanned = scans.get(key);
        if (scanned === undefined) {
          scanned = new WeakMap();
          scans.set(key, scanned);
        }
        readStory({ source, scanned, compiled, blocks, shown });
      }
      const unchanged =
        blocks.length === last.blocks.length &&
        blocks.every((block, index) => block === last.blocks[index]);
      const matches = unchanged
        ? last.matches
        : blocks.flatMap((block) =>
            block.flatMap((match) =>
              match.address === null ? [] : [match.address],
            ),
          );
      last = { blocks, matches };
      return { shown, matches };
    },
  };
};
