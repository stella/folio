/** Marker counters shared by parsing and canonical model normalization. */
import { panic } from "better-result";

import type { Paragraph } from "../types/document";
import { bulletMarkerFontName, convertBulletToUnicode } from "./bulletMarkers";
import { resolveListTemplate } from "./listTemplate";
import type { NumberingMap } from "./numberingParser";
import { isNumberingReference } from "./numberingReference";

export type PreviousListState = {
  abstractNumId: number | null;
  fromStyle: boolean;
  numId: number | null;
};

export type ComputeListMarkerOptions = {
  numbering: NumberingMap | null;
  listCounters: Map<number, number[]>;
  abstractCounters: Map<number, number[]>;
  restartedNumIds: Set<number>;
  previousList: PreviousListState;
};

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
  restartedNumIds: "story",
  previousList: "story",
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
  readonly restartedNumIds: readonly number[];
  readonly previousList: Readonly<PreviousListState>;
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
    restartedNumIds: [...counters.restartedNumIds],
    previousList: { ...counters.previousList },
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
    restartedNumIds: new Set(snapshot.restartedNumIds),
    previousList: { ...snapshot.previousList },
  };
  Object.assign(counters, story);
};

/**
 * The counter arrays one paragraph advanced: the array bound to its list
 * instance and abstract list after it, and those bound to them before it.
 * Besides these the paragraph changes only the previous-list state and
 * whether its instance restarted.
 */
export type ListCounterAdvance = {
  readonly numId: number;
  readonly abstractNumId: number | null;
  readonly counters: number[];
  readonly listBefore: number[] | undefined;
  readonly abstractBefore: number[] | undefined;
};

/** What one paragraph's marker computation left in the story half of the counters. */
export type ListCounterStep = {
  readonly previousList: Readonly<PreviousListState>;
  readonly advanced:
    | {
        readonly numId: number;
        readonly abstractNumId: number | null;
        readonly cell: number;
        readonly values: readonly number[];
        readonly restarted: boolean;
      }
    | undefined;
};

/** The step a paragraph just took, `advance` being what its computation returned. */
export const listCounterStep = (
  counters: ComputeListMarkerOptions,
  advance: ListCounterAdvance | null,
): ListCounterStep => ({
  previousList: { ...counters.previousList },
  advanced:
    advance === null
      ? undefined
      : {
          numId: advance.numId,
          abstractNumId: advance.abstractNumId,
          cell: counterCellOf(advance.counters),
          values: [...advance.counters],
          restarted: counters.restartedNumIds.has(advance.numId),
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
  Object.assign(counters.previousList, step.previousList);
  const advanced = step.advanced;
  if (advanced === undefined) return null;
  const { numId, abstractNumId, cell, values, restarted } = advanced;
  const listBefore = counters.listCounters.get(numId);
  const abstractBefore =
    abstractNumId === null
      ? undefined
      : counters.abstractCounters.get(abstractNumId);
  // The computation advances the instance's array, the abstract list's, or
  // one it made for the instance's first paragraph.
  const array =
    [listBefore, abstractBefore].find(
      (bound) => bound !== undefined && counterCellOf(bound) === cell,
    ) ?? cellArray(cell, values);
  for (let level = 0; level < values.length; level += 1)
    array[level] = values[level] ?? Number.NaN;
  counters.listCounters.set(numId, array);
  if (abstractNumId !== null)
    counters.abstractCounters.set(abstractNumId, array);
  if (restarted) counters.restartedNumIds.add(numId);
  return { numId, abstractNumId, counters: array, listBefore, abstractBefore };
};

export const computeListMarker = (
  paragraph: Paragraph,
  {
    numbering,
    listCounters,
    abstractCounters,
    restartedNumIds,
    previousList,
  }: ComputeListMarkerOptions,
): ListCounterAdvance | null => {
  const listRendering = paragraph.listRendering;
  if (!listRendering || !numbering) {
    previousList.abstractNumId = null;
    previousList.fromStyle = false;
    previousList.numId = null;
    return null;
  }

  const { numId, level } = listRendering;
  if (!isNumberingReference(numId)) {
    previousList.abstractNumId = null;
    previousList.fromStyle = false;
    previousList.numId = null;
    return null;
  }

  const listBefore = listCounters.get(numId);
  const firstEncounter = listBefore === undefined;
  let counters = listBefore ?? cellArray(numId, UNSET_COUNTERS);
  if (firstEncounter) listCounters.set(numId, counters);

  const abstractNumId = numbering.getAbstractNumId(numId);
  const styleNumbering =
    paragraph.formatting?.numPr?.kind !== "reference" &&
    paragraph.formatting?.numPrFromStyle?.kind === "reference";
  const resumesRestartedInstance =
    firstEncounter &&
    listRendering.startOverride === undefined &&
    abstractNumId !== null &&
    previousList.abstractNumId === abstractNumId &&
    previousList.numId !== null &&
    previousList.numId !== numId &&
    restartedNumIds.has(previousList.numId);
  const resumesStyleInstance =
    firstEncounter &&
    !styleNumbering &&
    listRendering.startOverride === undefined &&
    abstractNumId !== null &&
    previousList.abstractNumId === abstractNumId &&
    previousList.fromStyle === true;
  if (
    abstractNumId !== null &&
    (styleNumbering || resumesRestartedInstance || resumesStyleInstance)
  ) {
    const latestAbstractCounters = abstractCounters.get(abstractNumId);
    if (latestAbstractCounters) {
      // A paragraph whose numbering comes only from its style resumes the
      // latest compatible list instance. A fresh w:num with a startOverride
      // can restart the sequence before later paragraphs fall
      // back to the style's original w:num: the style continues the attachment
      // sequence instead of reviving its stale counters from earlier content.
      counters = latestAbstractCounters;
      listCounters.set(numId, counters);
    }
  }
  if (
    listRendering.startOverride !== undefined ||
    resumesRestartedInstance ||
    (previousList.numId === numId && restartedNumIds.has(numId))
  ) {
    restartedNumIds.add(numId);
  }
  if (abstractNumId !== null && level > 0) {
    const latestAbstractCounters = abstractCounters.get(abstractNumId);
    const missingParentCounters = counters.slice(0, level).every(Number.isNaN);
    if (missingParentCounters) {
      for (let i = 0; i < level; i += 1) {
        const latestCounter = latestAbstractCounters?.[i];
        counters[i] =
          latestCounter !== undefined && !Number.isNaN(latestCounter)
            ? latestCounter
            : (numbering.getLevel(numId, i)?.start ?? 1);
      }
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

  const abstractBefore =
    abstractNumId === null ? undefined : abstractCounters.get(abstractNumId);
  if (abstractNumId !== null) {
    abstractCounters.set(abstractNumId, counters);
  }
  const advance: ListCounterAdvance = {
    numId,
    abstractNumId,
    counters,
    listBefore,
    abstractBefore,
  };
  previousList.abstractNumId = abstractNumId;
  previousList.fromStyle = Boolean(styleNumbering);
  previousList.numId = numId;

  const pattern = listRendering.marker;

  if (listRendering.isBullet) {
    listRendering.marker = convertBulletToUnicode(
      pattern || "",
      bulletMarkerFontName(listRendering.markerFormatting),
    );
    previousList.abstractNumId = null;
    previousList.fromStyle = false;
    previousList.numId = null;
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
