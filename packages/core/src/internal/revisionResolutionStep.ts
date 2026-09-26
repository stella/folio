import { panic } from "better-result";
import type { Node as PMNode, Schema } from "prosemirror-model";
import { Step, StepMap, StepResult, type Mappable } from "prosemirror-transform";

import { indexedPositionMap } from "./revisionResolutionTracking";

type DeferredInverse = {
  doc: PMNode;
  steps: readonly Step[];
};

type RevisionResolutionStepOptions = {
  beforeDoc: PMNode | null;
  afterDoc: PMNode | null;
  steps: readonly Step[];
  positionMap: StepMap;
  deferredInverse?: DeferredInverse;
};

const mapThrough = (position: number, assoc: number, maps: readonly Mappable[]) => {
  let mapped = position;
  let deleted = false;
  let deletedBefore = false;
  let deletedAfter = false;
  let deletedAcross = false;
  for (const map of maps) {
    const result = map.mapResult(mapped, assoc);
    mapped = result.pos;
    deleted ||= result.deleted;
    deletedBefore ||= result.deletedBefore;
    deletedAfter ||= result.deletedAfter;
    deletedAcross ||= result.deletedAcross;
  }
  return {
    pos: mapped,
    deleted,
    deletedBefore,
    deletedAfter,
    deletedAcross,
  };
};

const chainedMap = (maps: readonly Mappable[]): Mappable => ({
  map: (position, assoc = 1) => mapThrough(position, assoc, maps).pos,
  mapResult: (position, assoc = 1) => mapThrough(position, assoc, maps),
});

const replaySteps = (doc: PMNode, steps: readonly Step[]): StepResult => {
  let current = doc;
  for (const step of steps) {
    const result = step.apply(current);
    if (!result.doc) {
      return result;
    }
    current = result.doc;
  }
  return StepResult.ok(current);
};

const isValidMapRanges = (value: unknown): value is number[] => {
  if (!Array.isArray(value) || value.length % 3 !== 0) {
    return false;
  }
  let previousEnd = 0;
  let delta = 0;
  for (let index = 0; index < value.length; index += 3) {
    const start = value.at(index);
    const oldSize = value.at(index + 1);
    const newSize = value.at(index + 2);
    if (
      !Number.isSafeInteger(start) ||
      !Number.isSafeInteger(oldSize) ||
      !Number.isSafeInteger(newSize) ||
      start < previousEnd ||
      oldSize < 0 ||
      newSize < 0
    ) {
      return false;
    }
    const oldEnd = start + oldSize;
    const newStart = start + delta;
    const newEnd = newStart + newSize;
    if (
      !Number.isSafeInteger(oldEnd) ||
      !Number.isSafeInteger(newStart) ||
      !Number.isSafeInteger(newEnd) ||
      !Number.isSafeInteger(newEnd - oldEnd) ||
      newStart < 0
    ) {
      return false;
    }
    previousEnd = oldEnd;
    delta = newEnd - oldEnd;
  }
  return true;
};

type IndexedMapChange = {
  oldStart: number;
  oldEnd: number;
  newStart: number;
  newEnd: number;
};

const indexedStepMapResult = (map: StepMap) => {
  const changes: IndexedMapChange[] = [];
  map.forEach((oldStart, oldEnd, newStart, newEnd) => {
    changes.push({ oldStart, oldEnd, newStart, newEnd });
  });
  return (position: number, assoc: number) => {
    let low = 0;
    let high = changes.length;
    while (low < high) {
      const middle = Math.floor((low + high) / 2);
      const change = changes[middle];
      if (change && change.oldEnd < position) low = middle + 1;
      else high = middle;
    }
    const change = changes[low];
    if (change && change.oldStart <= position) {
      const oldSize = change.oldEnd - change.oldStart;
      let side = assoc;
      if (oldSize > 0 && position === change.oldStart) side = -1;
      else if (oldSize > 0 && position === change.oldEnd) side = 1;
      return {
        pos: side < 0 ? change.newStart : change.newEnd,
        deleted: assoc < 0 ? position !== change.oldStart : position !== change.oldEnd,
      };
    }
    const previous = changes[low - 1];
    return { pos: position + (previous ? previous.newEnd - previous.oldEnd : 0), deleted: false };
  };
};

/** Compose by retaining precisely the source positions that survive both maps. */
export const composeRevisionResolutionMaps = (first: StepMap, second: StepMap): StepMap => {
  const firstResult = indexedStepMapResult(first);
  const secondResult = indexedStepMapResult(second);
  const inverseFirst = indexedPositionMap(first.invert());
  const firstBreakpoints = [0];
  const secondBreakpoints: number[] = [];
  let lastSecondBoundary: number | undefined;
  const projectSecondBoundary = (position: number) => {
    if (position === lastSecondBoundary) return;
    secondBreakpoints.push(inverseFirst(position, -1), inverseFirst(position, 1));
    lastSecondBoundary = position;
  };
  first.forEach((oldStart, oldEnd) => {
    firstBreakpoints.push(oldStart, oldEnd);
  });
  second.forEach((oldStart, oldEnd) => {
    projectSecondBoundary(oldStart);
    projectSecondBoundary(oldEnd);
  });
  const breakpoints: number[] = [];
  let firstIndex = 0;
  let secondIndex = 0;
  while (firstIndex < firstBreakpoints.length || secondIndex < secondBreakpoints.length) {
    const firstPosition = firstBreakpoints[firstIndex];
    const secondPosition = secondBreakpoints[secondIndex];
    const takeFirst =
      secondPosition === undefined ||
      (firstPosition !== undefined && firstPosition <= secondPosition);
    const position = takeFirst ? firstPosition : secondPosition;
    if (takeFirst) firstIndex++;
    else secondIndex++;
    if (position !== undefined && breakpoints.at(-1) !== position) breakpoints.push(position);
  }
  const through = (position: number, assoc: number) => {
    const earlier = firstResult(position, assoc);
    const later = secondResult(earlier.pos, assoc);
    return { pos: later.pos, deleted: earlier.deleted || later.deleted };
  };
  const ranges: number[] = [];
  let oldEnd = 0;
  let newEnd = 0;
  const anchor = (position: number, mapped: number) => {
    if (position > oldEnd || mapped > newEnd) {
      ranges.push(oldEnd, position - oldEnd, mapped - newEnd);
    }
    oldEnd = position;
    newEnd = mapped;
  };
  for (let index = 0; index < breakpoints.length; index++) {
    const position = breakpoints[index];
    if (position === undefined) continue;
    const left = through(position, -1);
    const right = through(position, 1);
    if (!left.deleted && !right.deleted) {
      anchor(position, Math.min(left.pos, right.pos));
      anchor(position, Math.max(left.pos, right.pos));
    } else if (!left.deleted) {
      anchor(position, left.pos);
    } else if (!right.deleted) {
      anchor(position, right.pos);
    }
    const next = breakpoints[index + 1];
    if (next === undefined) break;
    const middle = (position + next) / 2;
    const surviving = through(middle, 1);
    if (!surviving.deleted) {
      const offset = surviving.pos - middle;
      anchor(position, position + offset);
      oldEnd = next;
      newEnd = next + offset;
    }
  }
  return new StepMap(ranges);
};

/** One transaction step for a bulk resolution, with a granular position map. */
export class RevisionResolutionStep extends Step {
  private materializedSteps: readonly Step[] | null;
  private readonly beforeDoc: PMNode | null;
  private readonly afterDoc: PMNode | null;
  private readonly positionMap: StepMap;
  private readonly deferredInverse: DeferredInverse | null;

  constructor({
    beforeDoc,
    afterDoc,
    steps,
    positionMap,
    deferredInverse,
  }: RevisionResolutionStepOptions) {
    super();
    this.beforeDoc = beforeDoc;
    this.afterDoc = afterDoc;
    this.positionMap = positionMap;
    this.deferredInverse = deferredInverse ?? null;
    this.materializedSteps = deferredInverse ? null : steps;
  }

  private steps(): readonly Step[] {
    if (this.materializedSteps) {
      return this.materializedSteps;
    }
    const deferred = this.deferredInverse;
    if (!deferred) {
      return panic("Revision resolution has no inverse steps");
    }
    let doc = deferred.doc;
    const inverse: Step[] = [];
    for (const step of deferred.steps) {
      inverse.push(step.invert(doc));
      const result = step.apply(doc);
      if (!result.doc) {
        return panic(result.failed ?? "Revision resolution cannot invert a failed step");
      }
      doc = result.doc;
    }
    this.materializedSteps = inverse.reverse();
    return inverse;
  }

  override apply(doc: PMNode): StepResult {
    if (doc === this.beforeDoc && this.afterDoc) {
      return StepResult.ok(this.afterDoc);
    }
    return replaySteps(doc, this.steps());
  }

  override getMap(): StepMap {
    return this.positionMap;
  }

  override invert(doc: PMNode): RevisionResolutionStep {
    if (doc === this.beforeDoc && this.afterDoc) {
      return new RevisionResolutionStep({
        beforeDoc: this.afterDoc,
        afterDoc: doc,
        steps: [],
        positionMap: this.positionMap.invert(),
        deferredInverse: { doc, steps: this.steps() },
      });
    }
    let current = doc;
    const inverse: Step[] = [];
    for (const step of this.steps()) {
      inverse.push(step.invert(current));
      const result = step.apply(current);
      if (!result.doc) {
        return panic(result.failed ?? "Revision resolution cannot invert a failed step");
      }
      current = result.doc;
    }
    return new RevisionResolutionStep({
      beforeDoc: current,
      afterDoc: doc,
      steps: inverse.reverse(),
      positionMap: this.positionMap.invert(),
    });
  }

  override map(mapping: Mappable): RevisionResolutionStep | null {
    const original = this.steps();
    const mapped: Step[] = [];
    const inverseOriginalMaps: StepMap[] = [];
    const mappedMaps: StepMap[] = [];
    for (const step of original) {
      const through = chainedMap([...inverseOriginalMaps.toReversed(), mapping, ...mappedMaps]);
      const rebased = step.map(through);
      if (rebased) {
        mapped.push(rebased);
        mappedMaps.push(rebased.getMap());
      }
      inverseOriginalMaps.push(step.getMap().invert());
    }
    if (mapped.length === 0) {
      return null;
    }

    // Rebased edits can shrink or grow when remote edits overlap a removed
    // range. Derive the map from the surviving replay steps, just as a regular
    // transaction does, instead of dropping the entire resolution.
    let positionMap = StepMap.empty;
    for (const step of mapped) {
      positionMap = composeRevisionResolutionMaps(positionMap, step.getMap());
    }
    return new RevisionResolutionStep({
      beforeDoc: null,
      afterDoc: null,
      steps: mapped,
      positionMap,
    });
  }

  override toJSON(): unknown {
    const mapRanges: number[] = [];
    this.positionMap.forEach((oldStart, oldEnd, newStart, newEnd) => {
      mapRanges.push(oldStart, oldEnd - oldStart, newEnd - newStart);
    });
    return {
      stepType: "folioRevisionResolution",
      steps: this.steps().map((step) => step.toJSON()),
      mapRanges,
    };
  }

  static override fromJSON(schema: Schema, json: unknown): RevisionResolutionStep {
    if (typeof json !== "object" || json === null || !("steps" in json) || !("mapRanges" in json)) {
      throw new RangeError("Invalid revision resolution step JSON");
    }
    const { steps, mapRanges } = json;
    if (!Array.isArray(steps) || !isValidMapRanges(mapRanges)) {
      throw new RangeError("Invalid revision resolution step JSON");
    }
    return new RevisionResolutionStep({
      beforeDoc: null,
      afterDoc: null,
      steps: steps.map((step) => Step.fromJSON(schema, step)),
      positionMap: new StepMap(mapRanges),
    });
  }
}

Step.jsonID("folioRevisionResolution", RevisionResolutionStep);
