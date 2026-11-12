/**
 * What converting each top-level node of a document gave, kept for the next
 * conversion of the same editor's document.
 *
 * Converting a top-level node reads the node, its position, the conversion
 * constants and the state the nodes before it left: the list counters, the
 * last section's margins, how far the current section has come, and how many
 * block and content control ids were handed out. ProseMirror shares every
 * node an edit does not touch, so the next document starts and ends with the
 * same nodes. Where a kept node stands at the same position, after the same
 * ids and in the same state, it converts to the same blocks, which are handed
 * out again as they are. After the edit, the nodes are converted until the
 * state the edit left equals the state the kept node was converted in.
 *
 * A node whose blocks read other nodes is converted every time: a field may
 * show another paragraph's number, a text box anchor records which block
 * holds it, and a section-ending paragraph or page break reads its
 * neighbours.
 */

import type { Node as PMNode } from "prosemirror-model";

import { subtreeHolds } from "../../internal/subtreeHolds";
import type { FlowBlock } from "../../layout-engine/types";
import {
  cloneListCounterState,
  type ListCounterState,
} from "../../prosemirror/listMarker";
import type { ParagraphListMarker } from "./paragraphAttrs";

/** Margins, in twips, of the last section that authored them. */
export type SectionMarginsTwips = {
  readonly top: number;
  readonly bottom: number;
  readonly left: number;
  readonly right: number;
};

/** The state converting a top-level node starts from, besides ids. */
export type ConversionCheckpoint = {
  readonly final: ListCounterState;
  readonly original: ListCounterState;
  readonly margins: SectionMarginsTwips;
  /** Blocks the current section has projected so far. */
  readonly sectionBlocks: number;
};

/** One top-level node's conversion. */
export type TopLevelConversion = {
  readonly node: PMNode;
  readonly pos: number;
  /** The block ids converting the node handed out, in order. */
  readonly ids: readonly string[];
  /** Content control ids handed out before the node was converted. */
  readonly firstSdt: number;
  readonly sdts: number;
  readonly before: ConversionCheckpoint;
  /** The blocks the node projected, before the passes over the whole list. */
  readonly blocks: readonly FlowBlock[];
  /** For a node that converted to one paragraph block: the marker it had. */
  readonly paragraphMarker?: { readonly marker: ParagraphListMarker };
};

export type RecordedConversion = {
  readonly doc: PMNode;
  readonly nodes: readonly TopLevelConversion[];
  /** The state the last node left. */
  readonly end: ConversionCheckpoint;
  /** Above every block id the conversion handed out. */
  readonly freshFrom: number;
};

export type KeptBlockFlags = {
  /** Freeze every block a kept conversion keeps or hands out. */
  enabled?: boolean;
};

declare global {
  var __folioFreezeKeptBlocks: KeptBlockFlags | undefined;
}

/**
 * Whether kept conversions freeze their blocks. Tests set it, so code that
 * writes to a block handed out again fails where it writes instead of
 * changing what a later conversion hands out.
 */
export const freezesKeptBlocks = (): boolean =>
  globalThis.__folioFreezeKeptBlocks?.enabled === true;

/** Freeze every plain object and array `value` holds, and `value`. */
export const freezeKeptBlock = (value: unknown): void => {
  if (typeof value !== "object" || value === null || Object.isFrozen(value))
    return;
  const prototype: unknown = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== Array.prototype) return;
  Object.freeze(value);
  for (const nested of Object.values(value)) freezeKeptBlock(nested);
};

const readsOtherNodes = subtreeHolds((node) => {
  switch (node.type.name) {
    case "field":
    case "structuredField":
    case "textBoxAnchor":
    case "textBox":
    case "pageBreak":
    case "horizontalRule":
      return true;
    case "paragraph":
      return node.attrs["_sectionProperties"] != null;
    default:
      return false;
  }
});

/** Whether converting `node` reads anything besides itself and the state before it. */
export const convertsWithOtherNodes = (node: PMNode): boolean =>
  readsOtherNodes(node);

/** Two counter arrays read alike, `NaN` (no counter yet) included. */
const sameCounters = (
  left: readonly number[],
  right: readonly number[],
): boolean =>
  left.length === right.length &&
  left.every((value, index) => Object.is(value, right[index]));

const sameMembers = <T>(
  left: ReadonlySet<T> | undefined,
  right: ReadonlySet<T> | undefined,
): boolean => {
  const leftSize = left?.size ?? 0;
  if (leftSize !== (right?.size ?? 0)) return false;
  if (leftSize === 0) return true;
  for (const member of left ?? []) if (!right?.has(member)) return false;
  return true;
};

/**
 * Two counter states that number every paragraph after them alike. Advancing
 * a list writes one array that a numbering instance and its abstract
 * definition may both hold, so which arrays are shared is compared too.
 */
export const sameListCounterState = (
  left: ListCounterState,
  right: ListCounterState,
): boolean => {
  if (left === right) return true;
  if (
    left.counters.size !== right.counters.size ||
    left.abstractCounters.size !== right.abstractCounters.size ||
    left.lastAdvancedNumId !== right.lastAdvancedNumId ||
    !sameMembers(left.seenLevels, right.seenLevels) ||
    !sameMembers(left.restartedNumIds, right.restartedNumIds)
  )
    return false;
  const leftList = left.previousList;
  const rightList = right.previousList;
  if (
    (leftList?.abstractNumId ?? null) !== (rightList?.abstractNumId ?? null) ||
    (leftList?.fromStyle ?? false) !== (rightList?.fromStyle ?? false) ||
    (leftList?.numId ?? null) !== (rightList?.numId ?? null)
  )
    return false;
  const paired = new Map<readonly number[], readonly number[]>();
  const pairs = (
    leftArray: readonly number[],
    rightArray: readonly number[] | undefined,
  ): boolean => {
    if (rightArray === undefined) return false;
    const known = paired.get(leftArray);
    if (known !== undefined) return known === rightArray;
    if (!sameCounters(leftArray, rightArray)) return false;
    paired.set(leftArray, rightArray);
    return true;
  };
  const sameMap = (
    leftMap: ReadonlyMap<number, number[]>,
    rightMap: ReadonlyMap<number, number[]>,
  ): boolean => {
    const rightEntries = rightMap.entries();
    for (const [key, values] of leftMap) {
      const next = rightEntries.next();
      if (next.done === true) return false;
      const [rightKey, rightValues] = next.value;
      if (key !== rightKey || !pairs(values, rightValues)) return false;
    }
    return true;
  };
  return (
    sameMap(left.counters, right.counters) &&
    sameMap(left.abstractCounters, right.abstractCounters) &&
    // One array is never shared by two right-hand entries the left keeps apart.
    new Set(paired.values()).size === paired.size
  );
};

export const sameCheckpoint = (
  left: ConversionCheckpoint,
  right: ConversionCheckpoint,
): boolean =>
  left === right ||
  (left.sectionBlocks === right.sectionBlocks &&
    left.margins.top === right.margins.top &&
    left.margins.bottom === right.margins.bottom &&
    left.margins.left === right.margins.left &&
    left.margins.right === right.margins.right &&
    sameListCounterState(left.final, right.final) &&
    sameListCounterState(left.original, right.original));

export const takeCheckpoint = (
  final: ListCounterState,
  original: ListCounterState,
  margins: SectionMarginsTwips,
  sectionBlocks: number,
): ConversionCheckpoint => ({
  final: cloneListCounterState(final),
  original: cloneListCounterState(original),
  margins,
  sectionBlocks,
});

/**
 * Make `target` the state `source` holds, in place: the conversion options
 * hold `target`'s maps and sets themselves. Arrays `source` shares stay
 * shared in `target`, and `source` is never written.
 */
export const restoreListCounterState = (
  target: ListCounterState,
  source: ListCounterState,
): void => {
  const copies = new Map<readonly number[], number[]>();
  const copyOf = (values: readonly number[]): number[] => {
    const known = copies.get(values);
    if (known !== undefined) return known;
    const copy = [...values];
    copies.set(values, copy);
    return copy;
  };
  target.counters.clear();
  for (const [key, values] of source.counters)
    target.counters.set(key, copyOf(values));
  target.abstractCounters.clear();
  for (const [key, values] of source.abstractCounters)
    target.abstractCounters.set(key, copyOf(values));
  target.seenLevels.clear();
  for (const level of source.seenLevels) target.seenLevels.add(level);
  target.restartedNumIds = new Set(source.restartedNumIds);
  target.previousList = {
    ...(source.previousList ?? {
      abstractNumId: null,
      fromStyle: false,
      numId: null,
    }),
  };
  if (source.lastAdvancedNumId === undefined) delete target.lastAdvancedNumId;
  else target.lastAdvancedNumId = source.lastAdvancedNumId;
};
