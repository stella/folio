import type { StepMap } from "prosemirror-transform";

/** StepMap.map scans its ranges; index them once for many paragraph lookups. */
export const indexedPositionMap = (map: StepMap) => {
  const changes: { oldStart: number; oldEnd: number; newStart: number; newEnd: number }[] = [];
  map.forEach((oldStart, oldEnd, newStart, newEnd) => {
    changes.push({ oldStart, oldEnd, newStart, newEnd });
  });
  const mapResult = (position: number, assoc: 1 | -1) => {
    let low = 0;
    let high = changes.length;
    while (low < high) {
      const middle = Math.floor((low + high) / 2);
      const candidate = changes.at(middle);
      if (candidate && candidate.oldEnd < position) {
        low = middle + 1;
      } else {
        high = middle;
      }
    }
    const change = changes.at(low);
    if (change && change.oldStart <= position) {
      const oldSize = change.oldEnd - change.oldStart;
      let side = assoc;
      if (oldSize > 0 && position === change.oldStart) side = -1;
      else if (oldSize > 0 && position === change.oldEnd) side = 1;
      return {
        pos: side < 0 ? change.newStart : change.newEnd,
        deletedAcross: position > change.oldStart && position < change.oldEnd,
      };
    }
    const previous = low === 0 ? undefined : changes.at(low - 1);
    return {
      pos: position + (previous ? previous.newEnd - previous.oldEnd : 0),
      deletedAcross: false,
    };
  };
  return Object.assign((position: number, assoc: 1 | -1) => mapResult(position, assoc).pos, {
    mapResult,
  });
};
