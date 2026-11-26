/** Marker counters shared by parsing and canonical model normalization. */
import { panic } from "better-result";

import type { Paragraph } from "../types/document";
import { bulletMarkerFontName, convertBulletToUnicode } from "./bulletMarkers";
import { resolveListTemplate } from "./listTemplate";
import type { NumberingMap } from "./numberingParser";
import { isNumberingReference } from "./numberingReference";

export type ComputeListMarkerOptions = {
  numbering: NumberingMap | null;
  /** The counters each list instance counts on. */
  listCounters: Map<number, number[]>;
  /**
   * The counters of each list definition. Every instance of a definition
   * counts on these: a list goes on where the last paragraph of its
   * definition stands, whichever instance that paragraph named.
   */
  abstractCounters: Map<number, number[]>;
  /**
   * The levels of list instances a paragraph has used, as
   * {@link startedLevelKey}. An instance's start override sets a level's
   * number the first time the instance is used on that level, and never
   * again.
   */
  startedLevels: Set<number>;
};

const LEVELS_PER_INSTANCE = 16;

/** One level of one list instance, as a member of `startedLevels`. */
export const startedLevelKey = (numId: number, level: number): number =>
  numId * LEVELS_PER_INSTANCE + level;

/**
 * What each field of the marker counters depends on: the package alone, or
 * the paragraphs of the story before the current one. Total over the fields:
 * one added does not compile until it is placed, and the story half is what
 * a snapshot copies.
 */
const LIST_COUNTER_DEPENDENCY = {
  numbering: "package",
  listCounters: "story",
  abstractCounters: "story",
  startedLevels: "story",
} as const satisfies Record<
  keyof ComputeListMarkerOptions,
  "package" | "story"
>;
type CounterDependency = typeof LIST_COUNTER_DEPENDENCY;
type StoryCounterField = {
  [K in keyof CounterDependency]: CounterDependency[K] extends "story"
    ? K
    : never;
}[keyof CounterDependency];

/**
 * Counter arrays by the list instance that made them: an array is made when
 * its instance's first paragraph is met, once per story run, so the numId
 * names it the same way in every run over the same paragraphs.
 */
const counterCells = new WeakMap<readonly number[], number>();

const UNSET_COUNTERS: readonly number[] = Array.from<number>({
  length: 9,
}).fill(Number.NaN);

const cellArray = (cell: number, values: readonly number[]): number[] => {
  const array = [...values];
  counterCells.set(array, cell);
  return array;
};

/** The list instance whose first paragraph made `values`. */
export const counterCellOf = (values: readonly number[]): number =>
  counterCells.get(values) ?? panic("A list counter array has no cell.");

/**
 * The story half of the marker counters at one paragraph, copied. Counter
 * arrays are shared between list instances and their abstract list, and a
 * later paragraph advances the shared array, so the copy names each array by
 * its cell and keeps which entries share one.
 */
export type ListCounterSnapshot = {
  readonly cells: ReadonlyMap<number, readonly number[]>;
  readonly listCounters: readonly (readonly [number, number])[];
  readonly abstractCounters: readonly (readonly [number, number])[];
  readonly startedLevels: readonly number[];
} & Record<StoryCounterField, unknown>;

export const snapshotListCounters = (
  counters: ComputeListMarkerOptions,
): ListCounterSnapshot => {
  const cells = new Map<number, readonly number[]>();
  const cellOf = (values: number[]): number => {
    const cell = counterCellOf(values);
    if (!cells.has(cell)) cells.set(cell, [...values]);
    return cell;
  };
  return {
    listCounters: [...counters.listCounters].map(
      ([numId, values]) => [numId, cellOf(values)] as const,
    ),
    abstractCounters: [...counters.abstractCounters].map(
      ([abstractNumId, values]) => [abstractNumId, cellOf(values)] as const,
    ),
    cells,
    startedLevels: [...counters.startedLevels],
  };
};

/** Put the story half of `counters` back as the snapshot copied it. */
export const restoreListCounters = (
  counters: ComputeListMarkerOptions,
  snapshot: ListCounterSnapshot,
): void => {
  const arrays = new Map(
    [...snapshot.cells].map(([cell, values]) => [
      cell,
      cellArray(cell, values),
    ]),
  );
  const arrayOf = (cell: number): number[] =>
    arrays.get(cell) ?? panic("A list counter snapshot lost an array.");
  const story: Pick<ComputeListMarkerOptions, StoryCounterField> = {
    listCounters: new Map(
      snapshot.listCounters.map(([numId, cell]) => [numId, arrayOf(cell)]),
    ),
    abstractCounters: new Map(
      snapshot.abstractCounters.map(([abstractNumId, cell]) => [
        abstractNumId,
        arrayOf(cell),
      ]),
    ),
    startedLevels: new Set(snapshot.startedLevels),
  };
  Object.assign(counters, story);
};

/**
 * The counter arrays one paragraph advanced: the array bound to its list
 * instance and abstract list after it, and those bound to them before it.
 * Besides these the paragraph changes only which level of its instance has
 * been used.
 */
export type ListCounterAdvance = {
  readonly numId: number;
  readonly abstractNumId: number | null;
  /** The level the paragraph counted on. */
  readonly level: number;
  readonly counters: number[];
  readonly listBefore: number[] | undefined;
  readonly abstractBefore: number[] | undefined;
};

/** What one paragraph's marker computation left in the story half of the counters. */
export type ListCounterStep = {
  readonly advanced:
    | {
        readonly numId: number;
        readonly abstractNumId: number | null;
        readonly level: number;
        readonly cell: number;
        readonly values: readonly number[];
      }
    | undefined;
};

/** The step a paragraph just took, `advance` being what its computation returned. */
export const listCounterStep = (
  advance: ListCounterAdvance | null,
): ListCounterStep => ({
  advanced:
    advance === null
      ? undefined
      : {
          numId: advance.numId,
          abstractNumId: advance.abstractNumId,
          level: advance.level,
          cell: counterCellOf(advance.counters),
          values: [...advance.counters],
        },
});

/**
 * Take `step` again from the counters it was first taken from: the counters
 * end as the paragraph's computation left them, without the paragraph.
 */
export const replayListCounterStep = (
  counters: ComputeListMarkerOptions,
  step: ListCounterStep,
): ListCounterAdvance | null => {
  const advanced = step.advanced;
  if (advanced === undefined) return null;
  const { numId, abstractNumId, level, cell, values } = advanced;
  const listBefore = counters.listCounters.get(numId);
  const abstractBefore =
    abstractNumId === null
      ? undefined
      : counters.abstractCounters.get(abstractNumId);
  // The computation advances the abstract list's array, the instance's, or
  // one it made for the first paragraph of the list.
  const array =
    [abstractBefore, listBefore].find(
      (bound) => bound !== undefined && counterCellOf(bound) === cell,
    ) ?? cellArray(cell, values);
  for (let index = 0; index < values.length; index += 1)
    array[index] = values[index] ?? Number.NaN;
  counters.listCounters.set(numId, array);
  if (abstractNumId !== null)
    counters.abstractCounters.set(abstractNumId, array);
  counters.startedLevels.add(startedLevelKey(numId, level));
  return {
    numId,
    abstractNumId,
    level,
    counters: array,
    listBefore,
    abstractBefore,
  };
};

export const computeListMarker = (
  paragraph: Paragraph,
  {
    numbering,
    listCounters,
    abstractCounters,
    startedLevels,
  }: ComputeListMarkerOptions,
): ListCounterAdvance | null => {
  const listRendering = paragraph.listRendering;
  if (!listRendering || !numbering) {
    return null;
  }

  const { numId, level } = listRendering;
  if (!isNumberingReference(numId)) {
    return null;
  }

  const listBefore = listCounters.get(numId);
  const abstractNumId = numbering.getAbstractNumId(numId);
  const abstractBefore =
    abstractNumId === null ? undefined : abstractCounters.get(abstractNumId);
  // The instances of one definition are one list: a paragraph counts on from
  // the last paragraph of the definition, whichever instance that one named.
  // Only an instance without a definition counts alone.
  const counters =
    abstractBefore ?? listBefore ?? cellArray(numId, UNSET_COUNTERS);
  listCounters.set(numId, counters);

  // A start override takes effect once per level: on the first paragraph
  // that uses the instance there. Later paragraphs of the instance, also
  // after another instance was used in between, count on.
  const startedKey = startedLevelKey(numId, level);
  const startsLevel = !startedLevels.has(startedKey);
  startedLevels.add(startedKey);
  if (startsLevel && listRendering.startOverride !== undefined) {
    counters[level] = listRendering.startOverride - 1;
  }
  // A level used before any paragraph of the levels above counts those
  // levels at their start.
  if (
    abstractNumId !== null &&
    level > 0 &&
    counters.slice(0, level).every(Number.isNaN)
  ) {
    for (let i = 0; i < level; i += 1) {
      counters[i] = numbering.getLevel(numId, i)?.start ?? 1;
    }
  }

  if (Number.isNaN(counters[level])) {
    counters[level] = (numbering.getLevel(numId, level)?.start ?? 1) - 1;
  }
  counters[level] = (counters[level] ?? 0) + 1;

  for (let i = level + 1; i < counters.length; i += 1) {
    counters[i] = Number.NaN;
  }

  // A default LISTNUM field advances the counter one ilvl deeper
  // than its containing paragraph. Mirror the toFlowBlocks logic so the
  // marker substituted at parse time agrees with the renderer's counters —
  // otherwise a follow-up paragraph at that depth picks up the stale,
  // pre-substituted "(a)" instead of "(b)".
  const childAdvances = listRendering.implicitChildLevelAdvances ?? 0;
  if (childAdvances > 0 && level + 1 < counters.length) {
    const childCounter = counters[level + 1];
    counters[level + 1] =
      (childCounter === undefined || Number.isNaN(childCounter)
        ? 0
        : childCounter) + childAdvances;
  }

  if (abstractNumId !== null) {
    abstractCounters.set(abstractNumId, counters);
  }
  const advance: ListCounterAdvance = {
    numId,
    abstractNumId,
    level,
    counters,
    listBefore,
    abstractBefore,
  };

  const pattern = listRendering.marker;

  if (listRendering.isBullet) {
    listRendering.marker = convertBulletToUnicode(
      pattern || "",
      bulletMarkerFontName(listRendering.markerFormatting),
    );
    return advance;
  }

  const currentLevelInfo = numbering.getLevel(numId, level);
  listRendering.marker = resolveListTemplate({
    template: pattern,
    counters,
    levelFormats:
      listRendering.levelNumFmts ??
      counters.map((_counter, ilvl) => numbering.getLevel(numId, ilvl)?.numFmt),
    forceDecimal:
      currentLevelInfo?.isLgl === true || listRendering.isLegal === true,
  });
  return advance;
};
