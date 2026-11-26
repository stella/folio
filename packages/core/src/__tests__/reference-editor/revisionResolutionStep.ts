import { panic } from "better-result";
import type { Node as PMNode, Schema } from "prosemirror-model";
import {
  type Mappable,
  Step,
  StepMap,
  StepResult,
} from "prosemirror-transform";

import { composeRevisionResolutionMaps } from "./revisionResolutionMaps";

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

const mapThrough = (
  position: number,
  assoc: number,
  maps: readonly Mappable[],
) => {
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
        return panic(
          result.failed ?? "Revision resolution cannot invert a failed step",
        );
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
        return panic(
          result.failed ?? "Revision resolution cannot invert a failed step",
        );
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
      const through = chainedMap([
        ...inverseOriginalMaps.toReversed(),
        mapping,
        ...mappedMaps,
      ]);
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

  static override fromJSON(
    schema: Schema,
    json: unknown,
  ): RevisionResolutionStep {
    if (
      typeof json !== "object" ||
      json === null ||
      !("steps" in json) ||
      !("mapRanges" in json)
    ) {
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
