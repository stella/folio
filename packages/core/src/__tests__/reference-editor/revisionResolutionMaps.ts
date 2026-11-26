import { StepMap } from "prosemirror-transform";

import { indexedPositionMap } from "../../internal/indexedPositionMap";

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
        deleted:
          assoc < 0 ? position !== change.oldStart : position !== change.oldEnd,
      };
    }
    const previous = changes[low - 1];
    return {
      pos: position + (previous ? previous.newEnd - previous.oldEnd : 0),
      deleted: false,
    };
  };
};

/** Compose by retaining precisely the source positions that survive both maps. */
export const composeRevisionResolutionMaps = (
  first: StepMap,
  second: StepMap,
): StepMap => {
  const firstResult = indexedStepMapResult(first);
  const secondResult = indexedStepMapResult(second);
  const inverseFirst = indexedPositionMap(first.invert());
  const firstBreakpoints = [0];
  const secondBreakpoints: number[] = [];
  let lastSecondBoundary: number | undefined;
  const projectSecondBoundary = (position: number) => {
    if (position === lastSecondBoundary) return;
    secondBreakpoints.push(
      inverseFirst(position, -1),
      inverseFirst(position, 1),
    );
    lastSecondBoundary = position;
  };
  first.forEach((oldStart, oldEnd) => {
    firstBreakpoints.push(oldStart, oldEnd);
  });
  second.forEach((oldStart, oldEnd) => {
    projectSecondBoundary(oldStart);
    projectSecondBoundary(oldEnd);
  });
  // Boundaries inside content `first` inserted all project onto that
  // insertion's ends, so the projections repeat out of order.
  secondBreakpoints.sort((left, right) => left - right);
  const breakpoints: number[] = [];
  let firstIndex = 0;
  let secondIndex = 0;
  while (
    firstIndex < firstBreakpoints.length ||
    secondIndex < secondBreakpoints.length
  ) {
    const firstPosition = firstBreakpoints[firstIndex];
    const secondPosition = secondBreakpoints[secondIndex];
    const takeFirst =
      secondPosition === undefined ||
      (firstPosition !== undefined && firstPosition <= secondPosition);
    const position = takeFirst ? firstPosition : secondPosition;
    if (takeFirst) firstIndex++;
    else secondIndex++;
    if (position !== undefined && breakpoints.at(-1) !== position)
      breakpoints.push(position);
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
