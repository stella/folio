/**
 * Edits over a paragraph's content in its logical offset space.
 *
 * Every function here is pure: it returns new arrays and records along the
 * path it changed and hands back every child it did not change as the same
 * object, so an untouched run is `===` to its input.
 */

import { panic } from "better-result";

import type { ParagraphContent, Run, RunContent, TextFormatting } from "../model/document";
import { structurallyEqual } from "./equality";
import {
  asParagraphContent,
  type Gap,
  type InlineNode,
  isParagraphContent,
  mergeLists,
  partitionContent,
  runGaps,
  spanningRecords,
  zeroWidthLeavesAt,
} from "./leaves";
import {
  childrenOf,
  contentWidth,
  inlineWidth,
  isInlineContainer,
  isRemovedRevision,
  runContentWidth,
  runWidth,
  withChildren,
} from "./offsets";
import { applyFormattingPatch, priorValues } from "./patch";
import { identitySlots } from "./slots";
import {
  EMPTY_PROPERTY_SETS,
  type EmptyPropertySet,
  type FormattingPatch,
  INHERIT_RUN_PROPS,
  type InlineSlice,
  type InsertedRunProps,
  type NewIds,
  type RunPropsPatch,
} from "./types";

/** A run with the property set it states, or none. */
const withRunFormatting = (run: Run, formatting: TextFormatting | undefined): Run => {
  const next: Run = { ...run };
  if (formatting === undefined) {
    delete next.formatting;
  } else {
    next.formatting = formatting;
  }
  return next;
};

const withRunContent = (run: Run, content: RunContent[]): Run => ({ ...run, content });

/** Two property sets state the same thing; an absent set states nothing. */
export const sameRunFormatting = (
  left: TextFormatting | undefined,
  right: TextFormatting | undefined,
): boolean => structurallyEqual(left ?? {}, right ?? {});

/** How an inverse spells a property set it gives back: absent, or an empty set. */
export const emptySetSpelling = (formatting: object | undefined): EmptyPropertySet =>
  formatting !== undefined && Object.keys(formatting).length === 0
    ? EMPTY_PROPERTY_SETS.KEEP
    : EMPTY_PROPERTY_SETS.OMIT;

/** A patched property set, spelled as `whenEmpty` says once it has no keys left. */
export const patchedSet = <Formatting extends object>(
  base: Formatting | undefined,
  patch: FormattingPatch<Formatting>,
  whenEmpty: EmptyPropertySet | undefined,
): Partial<Formatting> | undefined =>
  applyFormattingPatch(base, patch) ?? (whenEmpty === EMPTY_PROPERTY_SETS.KEEP ? {} : undefined);

const newTextRun = (text: string, formatting: TextFormatting | undefined): Run => {
  const content: RunContent[] = [{ type: "text", text }];
  return formatting === undefined || Object.keys(formatting).length === 0
    ? { type: "run", content }
    : { type: "run", formatting, content };
};

// ---------------------------------------------------------------------------
// Delete and insert
// ---------------------------------------------------------------------------

/**
 * Remove the leaves between two gaps. A record the range passes through keeps
 * its two ends as one record; one left with nothing goes. Returns the new
 * content and the slice removed, with its cut ends.
 */
export const deleteBetween = (
  items: readonly ParagraphContent[],
  from: Gap,
  to: Gap,
): { content: ParagraphContent[]; removed: InlineSlice } => {
  const gaps = [from, to];
  const [before = [], middle = [], after = []] = partitionContent(items, gaps);
  const across = spanningRecords(items, gaps, 0, 2).length;
  const content =
    mergeLists(before, after, across) ?? panic("The two ends of a cut record always merge.");
  return {
    content: asParagraphContent(content),
    removed: {
      content: asParagraphContent(middle),
      openStart: spanningRecords(items, gaps, 0, 1).length,
      openEnd: spanningRecords(items, gaps, 1, 2).length,
    },
  };
};

/**
 * Cut the content at a gap and put a slice in, merging its open ends with the
 * halves; `undefined` when an open end does not match the record it meets.
 */
export const insertSliceAt = (
  items: readonly ParagraphContent[],
  at: Gap,
  slice: InlineSlice,
): ParagraphContent[] | undefined => {
  const [before = [], after = []] = partitionContent(items, [at]);
  const withStart = mergeLists(before, slice.content, slice.openStart);
  // An open end continues the record it meets, so that record's ids stand.
  const whole =
    withStart === undefined
      ? undefined
      : mergeLists(withStart, after, slice.openEnd, { mode: "exact", identity: "second" });
  return whole === undefined ? undefined : asParagraphContent(whole);
};

/** The gap after inserted content: its trailing zero-width leaves come before it. */
export const gapAfterInserted = (at: Gap, inserted: readonly ParagraphContent[]): Gap => {
  const width = contentWidth(inserted);
  const trailing = zeroWidthLeavesAt(inserted, width).length;
  return {
    offset: at.offset + width,
    zeroWidthBefore: width === 0 ? at.zeroWidthBefore + trailing : trailing,
  };
};

/**
 * What undoes an insertion: removing the inserted leaves, then merging the
 * records the insertion cut back to how `before` had them, `join` levels
 * below the ones still running across the gap. `undefined` when the
 * insertion merged records that were separate before it: no deletion could
 * tell them apart again, so such an insertion is refused.
 */
export const insertionInverse = (
  before: readonly ParagraphContent[],
  after: readonly ParagraphContent[],
  start: Gap,
  end: Gap,
): { removed: InlineSlice; join: number } | undefined => {
  const { content: restored, removed } = deleteBetween(after, start, end);
  const cut = spanningRecords(before, [start], 0, 1).length;
  const across = spanningRecords(restored, [start], 0, 1).length;
  return across > cut ? undefined : { removed, join: cut - across };
};

// ---------------------------------------------------------------------------
// Split and join inline records
// ---------------------------------------------------------------------------

/** Cut the innermost `depth` records running across a gap in two; `undefined` when fewer do. */
export const splitAt = (
  items: readonly ParagraphContent[],
  at: Gap,
  depth: number,
): ParagraphContent[] | undefined => {
  const spine = spanningRecords(items, [at], 0, 1);
  if (depth > spine.length) {
    return undefined;
  }
  const [before = [], after = []] = partitionContent(items, [at]);
  const content =
    mergeLists(before, after, spine.length - depth) ??
    panic("The two ends of a cut record always merge.");
  return asParagraphContent(content);
};

/**
 * A merge's result and the ids it retired: the second record's ids of each
 * pair merged, outermost first, in the order a cut recreating them takes
 * new ids.
 */
export type Joined =
  | { kind: "joined"; content: ParagraphContent[]; retired: NewIds }
  | { kind: "notAlike" }
  | { kind: "sharedId" };

/** Two sets of new ids, each space's in order: `first`'s then `second`'s. */
export const concatIds = (first: NewIds, second: NewIds): NewIds => {
  const revision = [...(first.revision ?? []), ...(second.revision ?? [])];
  const control = [...(first.control ?? []), ...(second.control ?? [])];
  return {
    ...(revision.length > 0 ? { revision } : {}),
    ...(control.length > 0 ? { control } : {}),
  };
};

/** Whether a set of new ids names any. */
export const namesIds = (ids: NewIds): boolean =>
  (ids.revision?.length ?? 0) > 0 || (ids.control?.length ?? 0) > 0;

/**
 * Merge `left` and `right`, `depth` levels, retiring ids from level `from`
 * down. Two records carrying the same id cannot be merged: a cut could not
 * give them back two.
 */
const mergeRetiring = (
  left: readonly InlineNode[],
  right: readonly InlineNode[],
  depth: number,
  from: number,
): Joined => {
  const retired: { revision: number[]; control: number[] } = { revision: [], control: [] };
  let shared = false;
  const content = mergeLists(left, right, depth, {
    mode: "exact",
    onMerge: (first, second, level) => {
      if (level < from) return;
      const kept = identitySlots(first);
      for (const [index, slot] of identitySlots(second).entries()) {
        shared ||= kept[index]?.id === slot.id;
        retired[slot.space].push(slot.id);
      }
    },
  });
  if (content === undefined) return { kind: "notAlike" };
  if (shared) return { kind: "sharedId" };
  return {
    kind: "joined",
    content: asParagraphContent(content),
    retired: concatIds({}, retired),
  };
};

/** Merge the records meeting at a gap, `depth` levels below the ones already running across it. */
export const joinAt = (items: readonly ParagraphContent[], at: Gap, depth: number): Joined => {
  const across = spanningRecords(items, [at], 0, 1).length;
  const [before = [], after = []] = partitionContent(items, [at]);
  return mergeRetiring(before, after, across + depth, across + 1);
};

/** A split paragraph's two halves and the records the cut ran through. */
export const cutAt = (
  items: readonly ParagraphContent[],
  at: Gap,
): { before: ParagraphContent[]; after: ParagraphContent[]; through: InlineNode[] } => {
  const [before = [], after = []] = partitionContent(items, [at]);
  return {
    before: asParagraphContent(before),
    after: asParagraphContent(after),
    through: spanningRecords(items, [at], 0, 1),
  };
};

/** Two paragraphs' content end to end, `depth` levels of the records meeting there merged. */
export const joinContent = (
  first: readonly ParagraphContent[],
  second: readonly ParagraphContent[],
  depth: number,
): Joined => mergeRetiring(first, second, depth, 1);

// ---------------------------------------------------------------------------
// Insert text
// ---------------------------------------------------------------------------

const splitRun = (run: Run, offset: number): { left: Run | undefined; right: Run | undefined } => {
  const left: RunContent[] = [];
  const right: RunContent[] = [];
  let position = 0;
  for (const child of run.content) {
    const width = runContentWidth(child);
    const start = position;
    position += width;
    if (position <= offset) {
      left.push(child);
    } else if (start >= offset) {
      right.push(child);
    } else if (child.type === "text") {
      left.push({ ...child, text: child.text.slice(0, offset - start) });
      right.push({ ...child, text: child.text.slice(offset - start) });
    }
  }
  return {
    left: left.length === 0 ? undefined : withRunContent(run, left),
    right: right.length === 0 ? undefined : withRunContent(run, right),
  };
};

type RunHit = {
  /** Child indices from the list down to the run. */
  path: number[];
  run: Run;
  /** Units of the run before the hit unit. */
  runOffset: number;
  insideRemovedRevision: boolean;
};

/** The run holding unit `unit`, or `undefined` when the unit is not run content. */
const runAtUnit = (
  items: readonly ParagraphContent[],
  unit: number,
  insideRemovedRevision: boolean,
): RunHit | undefined => {
  let position = 0;
  for (const [index, item] of items.entries()) {
    const width = inlineWidth(item);
    const start = position;
    position += width;
    if (unit < start || unit >= position) {
      continue;
    }
    if (item.type === "run") {
      return { path: [index], run: item, runOffset: unit - start, insideRemovedRevision };
    }
    if (isInlineContainer(item)) {
      const hit = runAtUnit(
        childrenOf(item),
        unit - start,
        insideRemovedRevision || isRemovedRevision(item),
      );
      return hit === undefined ? undefined : { ...hit, path: [index, ...hit.path] };
    }
    return undefined;
  }
  return undefined;
};

type InsertionHost = { path: number[]; run: Run; runOffset: number };

/**
 * The run inserted text joins: the one holding the unit before the position,
 * else the one holding the unit after it, skipping runs inside tracked
 * deletions and moved-away content.
 */
const insertionHost = (
  items: readonly ParagraphContent[],
  offset: number,
): InsertionHost | undefined => {
  const before = offset > 0 ? runAtUnit(items, offset - 1, false) : undefined;
  if (before !== undefined && !before.insideRemovedRevision) {
    return { path: before.path, run: before.run, runOffset: before.runOffset + 1 };
  }
  const after = runAtUnit(items, offset, false);
  if (after !== undefined && !after.insideRemovedRevision) {
    return { path: after.path, run: after.run, runOffset: after.runOffset };
  }
  return undefined;
};

/** Replace the child at `path` with `replacement`. */
const replaceAtPath = (
  items: readonly ParagraphContent[],
  path: readonly number[],
  replacement: readonly ParagraphContent[],
): ParagraphContent[] => {
  const [index, ...rest] = path;
  if (index === undefined) {
    return [...replacement];
  }
  const out = [...items];
  if (rest.length === 0) {
    out.splice(index, 1, ...replacement);
    return out;
  }
  const item = items[index];
  if (item === undefined || !isInlineContainer(item)) {
    return panic(`Inline path step ${index} does not name a container.`);
  }
  out[index] = withChildren(item, replaceAtPath(childrenOf(item), rest, replacement));
  return out;
};

/**
 * The run with text added at a unit offset, joining the text node there when
 * one does. An empty text node is a zero-width child of its own and is left
 * as it is.
 */
const runWithText = (run: Run, runOffset: number, text: string): Run => {
  let position = 0;
  let insertAt = 0;
  for (const [index, child] of run.content.entries()) {
    const width = runContentWidth(child);
    if (
      child.type === "text" &&
      width > 0 &&
      position <= runOffset &&
      runOffset <= position + width
    ) {
      const cut = runOffset - position;
      const content = [...run.content];
      content[index] = { ...child, text: child.text.slice(0, cut) + text + child.text.slice(cut) };
      return withRunContent(run, content);
    }
    position += width;
    if (position <= runOffset) {
      insertAt = index + 1;
    }
  }
  const content = [...run.content];
  content.splice(insertAt, 0, { type: "text", text });
  return withRunContent(run, content);
};

/**
 * Put a new run at a position that no run holds; `undefined` when the position
 * is inside tracked-deleted or moved-away content.
 */
const insertRunAtSlot = (
  items: readonly ParagraphContent[],
  offset: number,
  run: Run,
): ParagraphContent[] | undefined => {
  let position = 0;
  let insertAt = 0;
  for (const [index, item] of items.entries()) {
    const width = inlineWidth(item);
    const start = position;
    position += width;
    if (start < offset && offset < position && isInlineContainer(item)) {
      if (isRemovedRevision(item)) {
        return undefined;
      }
      const children = insertRunAtSlot(childrenOf(item), offset - start, run);
      if (children === undefined) {
        return undefined;
      }
      const out = [...items];
      out[index] = withChildren(item, children);
      return out;
    }
    if (position <= offset) {
      insertAt = index + 1;
    }
  }
  const out = [...items];
  out.splice(insertAt, 0, run);
  return out;
};

/**
 * Insert text at a position; `undefined` when the position is inside
 * tracked-deleted or moved-away content and no other run borders it.
 */
export const insertTextInContent = (
  items: readonly ParagraphContent[],
  offset: number,
  text: string,
  runProps: InsertedRunProps,
): ParagraphContent[] | undefined => {
  const host = insertionHost(items, offset);
  const formatting = runProps === INHERIT_RUN_PROPS ? undefined : runProps;
  if (host === undefined) {
    return insertRunAtSlot(items, offset, newTextRun(text, formatting));
  }
  if (formatting === undefined || sameRunFormatting(formatting, host.run.formatting)) {
    return replaceAtPath(items, host.path, [runWithText(host.run, host.runOffset, text)]);
  }
  const halves = splitRun(host.run, host.runOffset);
  const replacement: Run[] = [];
  if (halves.left !== undefined) replacement.push(halves.left);
  replacement.push(newTextRun(text, formatting));
  if (halves.right !== undefined) replacement.push(halves.right);
  return replaceAtPath(items, host.path, replacement);
};

// ---------------------------------------------------------------------------
// Run properties
// ---------------------------------------------------------------------------

const patchRunsIn = (
  nodes: readonly InlineNode[],
  patch: RunPropsPatch,
  whenEmpty: EmptyPropertySet | undefined,
  prior: Map<Run, TextFormatting | undefined>,
): InlineNode[] => {
  const out: InlineNode[] = [];
  for (const node of nodes) {
    if (node.type === "run") {
      const formatting = patchedSet(node.formatting, patch, whenEmpty);
      if (runWidth(node) === 0 || sameRunFormatting(formatting, node.formatting)) {
        out.push(node);
        continue;
      }
      const run = withRunFormatting(node, formatting);
      prior.set(run, node.formatting);
      out.push(run);
      continue;
    }
    if (isParagraphContent(node) && isInlineContainer(node)) {
      const children = childrenOf(node);
      const next = patchRunsIn(children, patch, whenEmpty, prior);
      const same = next.every((child, index) => child === children[index]);
      out.push(same ? node : withChildren(node, asParagraphContent(next)));
      continue;
    }
    out.push(node);
  }
  return out;
};

const collectRuns = (nodes: readonly InlineNode[], out: Run[]): void => {
  for (const node of nodes) {
    if (node.type === "run") {
      if (runWidth(node) > 0) out.push(node);
      continue;
    }
    if (isParagraphContent(node) && isInlineContainer(node)) {
      collectRuns(childrenOf(node), out);
    }
  }
};

/** The runs a patch between two gaps reaches, cut at the gaps: those with units there. */
export const runsBetween = (items: readonly ParagraphContent[], from: Gap, to: Gap): Run[] => {
  const [, middle = []] = partitionContent(items, [from, to]);
  const out: Run[] = [];
  collectRuns(middle, out);
  return out;
};

/** A patch over a stretch of runs, the shape of a `setRunProps` operation. */
export type RunPatchSpan = {
  from: Gap;
  to: Gap;
  patch: RunPropsPatch;
  whenEmpty: EmptyPropertySet;
};

/**
 * Patch the run properties of every run between two gaps, cutting runs at the
 * gaps. A run the patch leaves as it was is not cut. `undefined` when no run
 * changes.
 *
 * `restoring` gives the changed runs back the values the patch replaced, one
 * patch per stretch of adjacent runs that had the same values. A run the
 * patch left alone ends a stretch, so a restoring patch never reaches it.
 */
export const patchRunsBetween = (
  items: readonly ParagraphContent[],
  from: Gap,
  to: Gap,
  patch: RunPropsPatch,
  whenEmpty: EmptyPropertySet | undefined,
): { content: ParagraphContent[]; restoring: RunPatchSpan[] } | undefined => {
  const gaps = [from, to];
  const [before = [], middle = [], after = []] = partitionContent(items, gaps);
  const prior = new Map<Run, TextFormatting | undefined>();
  const patchedMiddle = patchRunsIn(middle, patch, whenEmpty, prior);
  if (prior.size === 0) {
    return undefined;
  }
  // A cut run the patch changed stays cut; containers and unchanged runs merge back.
  const alike = { mode: "asFarAsAlike" } as const;
  const head = mergeLists(before, patchedMiddle, spanningRecords(items, gaps, 0, 1).length, alike);
  const whole = mergeLists(head ?? [], after, spanningRecords(items, gaps, 1, 2).length, alike);
  const content = asParagraphContent(whole ?? []);

  const restoring: (RunPatchSpan & { key: string })[] = [];
  let stretchOpen = false;
  for (const { node, before: start, after: end } of runGaps(content)) {
    if (node.type !== "run") continue;
    if (!prior.has(node)) {
      if (runWidth(node) > 0) stretchOpen = false;
      continue;
    }
    const previous = prior.get(node);
    const restore = priorValues(previous, patch);
    const spelling = emptySetSpelling(previous);
    const key = JSON.stringify([restore, spelling]);
    const last = restoring.at(-1);
    if (stretchOpen && last !== undefined && last.key === key) {
      last.to = end;
    } else {
      restoring.push({ from: start, to: end, patch: restore, whenEmpty: spelling, key });
    }
    stretchOpen = true;
  }
  return {
    content,
    restoring: restoring.map(({ from: start, to: end, patch: restore, whenEmpty: spelling }) => ({
      from: start,
      to: end,
      patch: restore,
      whenEmpty: spelling,
    })),
  };
};
