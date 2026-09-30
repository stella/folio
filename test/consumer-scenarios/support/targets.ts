/**
 * Where edits collide: the blocks and offsets worth aiming at, and a picker
 * that aims there most of the time. A block's features come from what a
 * public reader exposes: the reviewer's blocks (list, table, note-reference
 * and run boundaries), its changes and comments, and `toDocument()` for what
 * blocks do not say (a paragraph carrying a section break, a field, a
 * hyperlink, a text box, a pending revision in a header or a note).
 *
 * `uniformPicker` draws exactly what the generators drew before targeting
 * existed, so a flow on the legacy generator replays its seed unchanged.
 */

import type { FolioDocumentStoryHandle } from "@stll/folio-core/server";

import type { openReviewer } from "./documents.ts";
import type { Random } from "./random.ts";
import { weightedChoice } from "./feature-coverage.ts";

type Reviewer = Awaited<ReturnType<typeof openReviewer>>;

/** Where a block lives, as the coverage ledger names it. */
export type StoryKind =
  | "body"
  | "tableCell"
  | "footnote"
  | "endnote"
  | "header"
  | "footer"
  | "textbox";

/** What makes a block (or an offset in it) an edge worth aiming at. */
export type Feature =
  | "sectionCarrier"
  | "listItem"
  | "tableEdge"
  | "inlineObject"
  | "pendingRevision"
  | "commentAnchor"
  | "field"
  | "hyperlink"
  | "noteReference"
  | "contentControl"
  | "formatBoundary"
  | "storyEdge"
  | "surrogateBoundary";

export const FEATURES: readonly Feature[] = [
  "sectionCarrier",
  "listItem",
  "tableEdge",
  "inlineObject",
  "pendingRevision",
  "commentAnchor",
  "field",
  "hyperlink",
  "noteReference",
  "contentControl",
  "formatBoundary",
  "storyEdge",
  "surrogateBoundary",
];

/** The fields of a reader's block the targeting reads. */
export type TargetBlock = {
  id: string;
  kind: string;
  text: string;
  listReference?: unknown;
  table?: unknown;
  previewRuns?: readonly { text: string }[];
  structuralBoundaries?: readonly { type: string; offset: number; length?: number }[];
};

type TableLocation = { tableIndex: number; rowIndex: number; cellIndex: number };

/** Every block's features, for one story of one reviewer state. */
export type FeatureIndex = {
  story: FolioDocumentStoryHandle;
  features: ReadonlyMap<string, ReadonlySet<Feature>>;
  /** Blocks a reader lists with the story that are paragraphs of a text box drawn in it. */
  inTextBox: ReadonlySet<string>;
};

// ---------------------------------------------------------------------------
// Features
// ---------------------------------------------------------------------------

const NODE_FEATURES: Readonly<Record<string, Feature>> = {
  shape: "inlineObject",
  drawing: "inlineObject",
  simpleField: "field",
  complexField: "field",
  fieldChar: "field",
  hyperlink: "hyperlink",
  commentRangeStart: "commentAnchor",
  commentRangeEnd: "commentAnchor",
  commentReference: "commentAnchor",
  insertion: "pendingRevision",
  deletion: "pendingRevision",
  moveFrom: "pendingRevision",
  moveTo: "pendingRevision",
  footnoteRef: "noteReference",
  endnoteRef: "noteReference",
  sdt: "contentControl",
};

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null;

/** Features found anywhere under `node`. */
const collectFeatures = (node: unknown, into: Set<Feature>): void => {
  if (Array.isArray(node)) {
    for (const item of node) collectFeatures(item, into);
    return;
  }
  if (!isRecord(node)) return;
  const feature = typeof node["type"] === "string" ? NODE_FEATURES[node["type"]] : undefined;
  if (feature) into.add(feature);
  if (node["pPrMark"] !== undefined) into.add("pendingRevision");
  if (Array.isArray(node["contentControls"]) && node["contentControls"].length > 0) {
    into.add("contentControl");
  }
  if (Array.isArray(node["propertyChanges"]) && node["propertyChanges"].length > 0) {
    into.add("pendingRevision");
  }
  for (const [key, value] of Object.entries(node)) {
    // Authored XML kept for replay is not structure.
    if (key === "rawXml" || key === "verbatimXml") continue;
    if (isRecord(value) || Array.isArray(value)) collectFeatures(value, into);
  }
};

/**
 * Every paragraph under `root` with a `paraId`, with the features its model
 * shows; `inTextBox` collects the ones inside a text box.
 */
export const paragraphFeatures = (
  root: unknown,
  into: Map<string, Set<Feature>>,
  inTextBox: Set<string>,
): void => {
  const visit = (node: unknown, boxed: boolean, controlled: boolean): void => {
    if (Array.isArray(node)) {
      for (const item of node) visit(item, boxed, controlled);
      return;
    }
    if (!isRecord(node)) return;
    if (node["type"] === "paragraph" && typeof node["paraId"] === "string") {
      if (boxed) inTextBox.add(node["paraId"].toUpperCase());
      const features = new Set<Feature>();
      collectFeatures(node["content"], features);
      if (controlled) features.add("contentControl");
      if (node["pPrMark"] !== undefined) features.add("pendingRevision");
      if (Array.isArray(node["propertyChanges"]) && node["propertyChanges"].length > 0) {
        features.add("pendingRevision");
      }
      if (node["sectionProperties"] !== undefined) features.add("sectionCarrier");
      into.set(node["paraId"].toUpperCase(), features);
    }
    const box = boxed || node["type"] === "shape" || node["type"] === "textBox";
    const inControl = controlled || node["type"] === "sdt";
    for (const value of Object.values(node)) {
      if (isRecord(value) || Array.isArray(value)) visit(value, box, inControl);
    }
  };
  visit(root, false, false);
};

/** The model content of one story. */
const storyContent = (reviewer: Reviewer, story: FolioDocumentStoryHandle): unknown => {
  const pkg = reviewer.toDocument().package;
  switch (story.type) {
    case "main":
      return pkg.document.content;
    case "header":
      return pkg.headers?.get(story.relationshipId)?.content;
    case "footer":
      return pkg.footers?.get(story.relationshipId)?.content;
    case "footnote":
      return pkg.footnotes?.find((note) => note.id === story.noteId)?.content;
    case "endnote":
      return pkg.endnotes?.find((note) => note.id === story.noteId)?.content;
  }
};

const ASTRAL = /[\u{10000}-\u{10FFFF}]/u;

/** The blocks of `story`, as a reader lists them. */
export const blocksOfStory = (
  reviewer: Reviewer,
  story: FolioDocumentStoryHandle = { type: "main" },
): TargetBlock[] =>
  story.type === "main"
    ? (reviewer.getContent() as TargetBlock[])
    : ((reviewer.snapshotStory(story)?.blocks ?? []) as TargetBlock[]);

const tableEdge = (block: TargetBlock, blocks: readonly TargetBlock[]): boolean => {
  const at = block.table as TableLocation | undefined;
  if (!at) return false;
  const same = blocks
    .map((candidate) => candidate.table as TableLocation | undefined)
    .filter((location) => location?.tableIndex === at.tableIndex) as TableLocation[];
  const rows = same.map((location) => location.rowIndex);
  const cells = same
    .filter((location) => location.rowIndex === at.rowIndex)
    .map((location) => location.cellIndex);
  return (
    at.rowIndex === Math.min(...rows) ||
    at.rowIndex === Math.max(...rows) ||
    at.cellIndex === Math.min(...cells) ||
    at.cellIndex === Math.max(...cells)
  );
};

/** The features of every block of `story` in the reviewer's current state. */
export const featureIndex = (
  reviewer: Reviewer,
  story: FolioDocumentStoryHandle = { type: "main" },
): FeatureIndex => {
  const blocks = blocksOfStory(reviewer, story);
  const model = new Map<string, Set<Feature>>();
  const boxed = new Set<string>();
  paragraphFeatures(storyContent(reviewer, story), model, boxed);
  const pending = new Set<string>();
  const anchored = new Set<string>();
  if (story.type === "main") {
    for (const change of reviewer.getChanges()) if (change.blockId) pending.add(change.blockId);
    for (const comment of reviewer.getComments())
      if (comment.blockId) anchored.add(comment.blockId);
  }
  // A block between the paragraphs a comment range starts and ends in sits inside it.
  let insideComment = false;
  const features = new Map<string, Set<Feature>>();
  blocks.forEach((block, index) => {
    const found = new Set<Feature>(model.get(block.id.toUpperCase()) ?? []);
    if (pending.has(block.id)) found.add("pendingRevision");
    if (anchored.has(block.id)) found.add("commentAnchor");
    const opensComment = found.has("commentAnchor");
    if (insideComment) found.add("commentAnchor");
    if (opensComment) insideComment = !insideComment;
    if (block.listReference !== undefined || block.kind === "listItem") found.add("listItem");
    if (tableEdge(block, blocks)) found.add("tableEdge");
    if (block.structuralBoundaries?.some((boundary) => boundary.type === "noteReference")) {
      found.add("noteReference");
    }
    if ((block.previewRuns?.length ?? 0) > 1) found.add("formatBoundary");
    if (index === 0 || index === blocks.length - 1) found.add("storyEdge");
    if (ASTRAL.test(block.text)) found.add("surrogateBoundary");
    features.set(block.id, found);
  });
  const inTextBox = blocks.filter((block) => boxed.has(block.id.toUpperCase()));
  return { story, features, inTextBox: new Set(inTextBox.map(({ id }) => id)) };
};

// ---------------------------------------------------------------------------
// Offsets
// ---------------------------------------------------------------------------

export type Span = { word: string; start: number };

export const wordsOf = (text: string): Span[] =>
  [...text.matchAll(/[\p{L}\p{N}]{3,}/gu)].map((match) => ({
    word: match[0],
    start: match.index,
  }));

/** The UTF-16 ranges of the user-perceived characters that hold an astral code point. */
export const astralGraphemes = (text: string): { start: number; end: number }[] =>
  [...new Intl.Segmenter("en", { granularity: "grapheme" }).segment(text)]
    .filter(({ segment }) => ASTRAL.test(segment))
    .map(({ segment, index }) => ({ start: index, end: index + segment.length }));

/**
 * The offsets of `block` where something ends and something else starts: run
 * (format) boundaries, note-reference markers, and astral characters.
 */
export const boundariesOf = (block: TargetBlock): number[] => {
  const points = new Set<number>();
  let offset = 0;
  const runs = block.previewRuns ?? [];
  if (runs.map((run) => run.text).join("") === block.text) {
    for (const run of runs.slice(0, -1)) {
      offset += run.text.length;
      points.add(offset);
    }
  }
  for (const boundary of block.structuralBoundaries ?? []) {
    points.add(boundary.offset);
    points.add(boundary.offset + (boundary.length ?? 0));
  }
  for (const { start, end } of astralGraphemes(block.text)) {
    points.add(start);
    points.add(end);
  }
  return [...points]
    .filter((point) => point > 0 && point < block.text.length)
    .sort((a, b) => a - b);
};

/** Whether `[start, end)` begins or ends right beside an astral character. */
export const touchesSurrogate = (text: string, start: number, end: number): boolean =>
  astralGraphemes(text).some(
    (grapheme) =>
      grapheme.start === start ||
      grapheme.end === start ||
      grapheme.start === end ||
      grapheme.end === end,
  );

/** Spans that start or end on a boundary: the word before or after it, or the marker itself. */
const edgeSpans = (block: TargetBlock): Span[] => {
  const words = wordsOf(block.text);
  const spans: Span[] = [];
  for (const point of boundariesOf(block)) {
    const before = words.findLast((word) => word.start + word.word.length <= point);
    const after = words.find((word) => word.start >= point);
    if (before)
      spans.push(before, { word: block.text.slice(before.start, point), start: before.start });
    if (after) {
      spans.push(after, {
        word: block.text.slice(point, after.start + after.word.length),
        start: point,
      });
    }
  }
  for (const boundary of block.structuralBoundaries ?? []) {
    if (boundary.type === "noteReference" && boundary.length) {
      spans.push({
        word: block.text.slice(boundary.offset, boundary.offset + boundary.length),
        start: boundary.offset,
      });
    }
  }
  for (const { start, end } of astralGraphemes(block.text)) {
    spans.push({ word: block.text.slice(start, end), start });
  }
  // A span of only spaces finds nothing a model would ask for.
  return spans.filter((span) => span.word.trim().length > 0);
};

// ---------------------------------------------------------------------------
// Pickers
// ---------------------------------------------------------------------------

/** How a generator chooses its block, word and split point. */
export type Picker = {
  block: <T extends TargetBlock>(candidates: readonly T[]) => T;
  /** A span of the block's text to find, replace, format or comment on. */
  span: (block: TargetBlock) => Span;
  /** Where to split a block that has at least two words. */
  split: (block: TargetBlock) => number;
};

/** The draws the generators made before targeting: uniform, and in the same order. */
export const uniformPicker = (random: Random): Picker => ({
  block: (candidates) => random.pick(candidates),
  span: (block) => random.pick(wordsOf(block.text)),
  split: (block) => random.pick(wordsOf(block.text).slice(1)).start,
});

export type BiasOptions = {
  index: FeatureIndex;
  /** Blocks earlier steps touched, most recent last. */
  recent: readonly string[];
  /** The share of picks that stay uniform. */
  uniform?: number;
  coverageWeight?: (block: TargetBlock) => number;
};

/**
 * Aim at edges: most picks go to a recently touched block or one with a
 * feature, and to a span or split point on a boundary; the rest stay uniform.
 */
export const biasedPicker = (random: Random, options: BiasOptions): Picker => {
  const uniform = uniformPicker(random);
  const share = options.uniform ?? 0.25;
  const recent = new Set(options.recent.slice(-6));
  const featured = (block: TargetBlock) =>
    (options.index.features.get(block.id)?.size ?? 0) > 0 || options.index.inTextBox.has(block.id);
  return {
    block: (candidates) => {
      if (random.chance(share)) return uniform.block(candidates);
      if (options.coverageWeight) {
        return weightedChoice(candidates, random, options.coverageWeight);
      }
      const touched = candidates.filter((block) => recent.has(block.id));
      if (touched.length > 0 && random.chance(0.35)) return random.pick(touched);
      const hot = candidates.filter(featured);
      return hot.length > 0 ? random.pick(hot) : uniform.block(candidates);
    },
    span: (block) => {
      const edges = edgeSpans(block);
      if (edges.length === 0 || random.chance(share)) return uniform.span(block);
      return random.pick(edges);
    },
    split: (block) => {
      const words = wordsOf(block.text);
      const firstWordEnd = (words[0]?.start ?? 0) + (words[0]?.word.length ?? 0);
      const points = boundariesOf(block).filter((point) => point >= firstWordEnd);
      if (points.length === 0 || random.chance(share)) return uniform.split(block);
      return random.pick(points);
    },
  };
};

/** The ledger's story for a block of the story `index` describes. */
export const storyKindOf = (index: FeatureIndex, block: TargetBlock | undefined): StoryKind => {
  if (block && index.inTextBox.has(block.id)) return "textbox";
  if (index.story.type !== "main") return index.story.type;
  return block?.table !== undefined ? "tableCell" : "body";
};
