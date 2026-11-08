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
 * The story half of the marker counters at one paragraph, copied. Counter
 * arrays are shared between a list instance and its abstract list, and a
 * later paragraph advances the shared array, so the copy names each array by
 * its index in `arrays` and keeps which entries share one.
 */
export type ListCounterSnapshot = {
  readonly arrays: readonly (readonly number[])[];
  readonly listCounters: readonly (readonly [number, number])[];
  readonly abstractCounters: readonly (readonly [number, number])[];
  readonly restartedNumIds: readonly number[];
  readonly previousList: Readonly<PreviousListState>;
} & Record<StoryCounterField, unknown>;

export const snapshotListCounters = (
  counters: ComputeListMarkerOptions,
): ListCounterSnapshot => {
  const indices = new Map<number[], number>();
  const arrays: (readonly number[])[] = [];
  const indexOf = (values: number[]): number => {
    const known = indices.get(values);
    if (known !== undefined) return known;
    indices.set(values, arrays.length);
    arrays.push([...values]);
    return arrays.length - 1;
  };
  return {
    listCounters: [...counters.listCounters].map(
      ([numId, values]) => [numId, indexOf(values)] as const,
    ),
    abstractCounters: [...counters.abstractCounters].map(
      ([abstractNumId, values]) => [abstractNumId, indexOf(values)] as const,
    ),
    arrays,
    restartedNumIds: [...counters.restartedNumIds],
    previousList: { ...counters.previousList },
  };
};

/** Put the story half of `counters` back as the snapshot copied it. */
export const restoreListCounters = (
  counters: ComputeListMarkerOptions,
  snapshot: ListCounterSnapshot,
): void => {
  const arrays = snapshot.arrays.map((values) => [...values]);
  const arrayAt = (index: number): number[] =>
    arrays[index] ?? panic("A list counter snapshot lost an array.");
  const story: Pick<ComputeListMarkerOptions, StoryCounterField> = {
    listCounters: new Map(
      snapshot.listCounters.map(([numId, index]) => [numId, arrayAt(index)]),
    ),
    abstractCounters: new Map(
      snapshot.abstractCounters.map(([abstractNumId, index]) => [
        abstractNumId,
        arrayAt(index),
      ]),
    ),
    restartedNumIds: new Set(snapshot.restartedNumIds),
    previousList: { ...snapshot.previousList },
  };
  Object.assign(counters, story);
};

const samePairs = (
  left: readonly (readonly [number, number])[],
  right: readonly (readonly [number, number])[],
) =>
  left.length === right.length &&
  left.every(
    ([key, index], at) => right[at]?.[0] === key && right[at]?.[1] === index,
  );

/** Whether two snapshots hold the same counters, shared the same way, in the same order. */
export const sameListCounters = (
  left: ListCounterSnapshot,
  right: ListCounterSnapshot,
): boolean =>
  left.previousList.abstractNumId === right.previousList.abstractNumId &&
  left.previousList.fromStyle === right.previousList.fromStyle &&
  left.previousList.numId === right.previousList.numId &&
  left.restartedNumIds.length === right.restartedNumIds.length &&
  left.restartedNumIds.every(
    (numId, at) => right.restartedNumIds[at] === numId,
  ) &&
  samePairs(left.listCounters, right.listCounters) &&
  samePairs(left.abstractCounters, right.abstractCounters) &&
  left.arrays.length === right.arrays.length &&
  left.arrays.every((values, at) => {
    const other = right.arrays[at];
    return (
      other !== undefined &&
      values.length === other.length &&
      values.every((value, index) => Object.is(value, other[index]))
    );
  });

export const computeListMarker = (
  paragraph: Paragraph,
  {
    numbering,
    listCounters,
    abstractCounters,
    restartedNumIds,
    previousList,
  }: ComputeListMarkerOptions,
): void => {
  const listRendering = paragraph.listRendering;
  if (!listRendering || !numbering) {
    previousList.abstractNumId = null;
    previousList.fromStyle = false;
    previousList.numId = null;
    return;
  }

  const { numId, level } = listRendering;
  if (!isNumberingReference(numId)) {
    previousList.abstractNumId = null;
    previousList.fromStyle = false;
    previousList.numId = null;
    return;
  }

  const firstEncounter = !listCounters.has(numId);
  if (firstEncounter) {
    listCounters.set(numId, Array.from<number>({ length: 9 }).fill(Number.NaN));
  }

  let counters = listCounters.get(numId);
  if (!counters) {
    return;
  }

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

  if (abstractNumId !== null) {
    abstractCounters.set(abstractNumId, counters);
  }
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
    return;
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
};
