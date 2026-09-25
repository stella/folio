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
  childrenOf,
  inlineWidth,
  isIdentifiedContainer,
  isInlineContainer,
  isOpeningMarker,
  isRemovedRevision,
  runContentWidth,
  runWidth,
  withChildren,
} from "./offsets";
import { applyFormattingPatch } from "./patch";
import { INHERIT_RUN_PROPS, type InsertedRunProps, type RunPropsPatch } from "./types";

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

const newTextRun = (text: string, formatting: TextFormatting | undefined): Run => {
  const content: RunContent[] = [{ type: "text", text }];
  return formatting === undefined || Object.keys(formatting).length === 0
    ? { type: "run", content }
    : { type: "run", formatting, content };
};

// ---------------------------------------------------------------------------
// Delete
// ---------------------------------------------------------------------------

const deleteInRun = (run: Run, from: number, to: number): Run | undefined => {
  const content: RunContent[] = [];
  let position = 0;
  for (const child of run.content) {
    const width = runContentWidth(child);
    const start = position;
    position += width;
    if (width === 0 || position <= from || start >= to) {
      content.push(child);
      continue;
    }
    if (child.type === "text") {
      const text = child.text.slice(0, Math.max(0, from - start)) + child.text.slice(to - start);
      if (text !== "") {
        content.push({ ...child, text });
      }
    }
  }
  const kept = withRunContent(run, content);
  return runWidth(kept) === 0 ? undefined : kept;
};

/**
 * Remove units `[from, to)` of a child list. Zero-width children stay; a run
 * or container the removal leaves without units or children goes.
 */
export const deleteInContent = (
  items: readonly ParagraphContent[],
  from: number,
  to: number,
): ParagraphContent[] => {
  const out: ParagraphContent[] = [];
  let position = 0;
  for (const item of items) {
    const width = inlineWidth(item);
    const start = position;
    position += width;
    if (width === 0 || position <= from || start >= to) {
      out.push(item);
      continue;
    }
    if (item.type === "run") {
      const kept = deleteInRun(item, from - start, to - start);
      if (kept !== undefined) {
        out.push(kept);
      }
      continue;
    }
    if (isInlineContainer(item)) {
      const children = deleteInContent(childrenOf(item), from - start, to - start);
      if (children.length > 0) {
        out.push(withChildren(item, children));
      }
    }
    // An atom inside the range goes.
  }
  return out;
};

// ---------------------------------------------------------------------------
// Split
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

type SplitContent = { left: ParagraphContent[]; right: ParagraphContent[] };

/**
 * Cut a child list at a position. A zero-width boundary exactly at the cut
 * that opens a range goes with what follows it; every other zero-width child
 * there stays with what precedes it. A run or container the cut passes through
 * becomes two, each keeping every field of the original.
 *
 * `undefined` when the cut passes through a tracked change or content control:
 * both halves would carry its one id.
 */
export const splitContent = (
  items: readonly ParagraphContent[],
  offset: number,
): SplitContent | undefined => {
  const left: ParagraphContent[] = [];
  const right: ParagraphContent[] = [];
  let position = 0;
  for (const item of items) {
    const width = inlineWidth(item);
    const start = position;
    position += width;
    if (width === 0) {
      const goesRight = start > offset || (start === offset && isOpeningMarker(item));
      (goesRight ? right : left).push(item);
      continue;
    }
    if (position <= offset) {
      left.push(item);
      continue;
    }
    if (start >= offset) {
      right.push(item);
      continue;
    }
    if (item.type === "run") {
      const halves = splitRun(item, offset - start);
      if (halves.left !== undefined) left.push(halves.left);
      if (halves.right !== undefined) right.push(halves.right);
      continue;
    }
    if (isInlineContainer(item)) {
      if (isIdentifiedContainer(item)) {
        return undefined;
      }
      const halves = splitContent(childrenOf(item), offset - start);
      if (halves === undefined) {
        return undefined;
      }
      left.push(withChildren(item, halves.left));
      right.push(withChildren(item, halves.right));
    }
    // An atom has width 1, so a cut can only fall before or after it.
  }
  return { left, right };
};

// ---------------------------------------------------------------------------
// Insert text
// ---------------------------------------------------------------------------

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

/** The run with text added at a unit offset, joining the text node there when one does. */
const runWithText = (run: Run, runOffset: number, text: string): Run => {
  let position = 0;
  let insertAt = 0;
  for (const [index, child] of run.content.entries()) {
    const width = runContentWidth(child);
    if (child.type === "text" && position <= runOffset && runOffset <= position + width) {
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

const patchRun = (run: Run, from: number, to: number, patch: RunPropsPatch): Run[] => {
  const patched = applyFormattingPatch(run.formatting, patch);
  if (sameRunFormatting(patched, run.formatting)) {
    return [run];
  }
  const width = runWidth(run);
  const pieces: Run[] = [];
  let rest: Run | undefined = run;
  let restStart = 0;
  if (from > 0) {
    const halves = splitRun(run, from);
    if (halves.left !== undefined) pieces.push(halves.left);
    rest = halves.right;
    restStart = from;
  }
  if (rest === undefined) {
    return pieces;
  }
  let tail: Run | undefined;
  if (to < width) {
    const halves = splitRun(rest, to - restStart);
    rest = halves.left;
    tail = halves.right;
  }
  if (rest !== undefined) pieces.push(withRunFormatting(rest, patched));
  if (tail !== undefined) pieces.push(tail);
  return pieces;
};

/**
 * Patch the run properties of every run unit in `[from, to)`, splitting runs at
 * the range ends. A run the patch leaves as it was is not split, and a list
 * nothing in changed is returned as the same array.
 */
export const patchRunsInContent = (
  items: readonly ParagraphContent[],
  from: number,
  to: number,
  patch: RunPropsPatch,
): readonly ParagraphContent[] => {
  const out: ParagraphContent[] = [];
  let changed = false;
  let position = 0;
  for (const item of items) {
    const width = inlineWidth(item);
    const start = position;
    position += width;
    if (width === 0 || position <= from || start >= to) {
      out.push(item);
      continue;
    }
    if (item.type === "run") {
      const pieces = patchRun(item, Math.max(0, from - start), Math.min(width, to - start), patch);
      changed ||= pieces.length !== 1 || pieces[0] !== item;
      out.push(...pieces);
      continue;
    }
    if (isInlineContainer(item)) {
      const children = childrenOf(item);
      const patched = patchRunsInContent(children, from - start, to - start, patch);
      if (patched === children) {
        out.push(item);
      } else {
        changed = true;
        out.push(withChildren(item, patched));
      }
      continue;
    }
    // Fields, equations and other atoms carry no run properties of their own here.
    out.push(item);
  }
  return changed ? out : items;
};
