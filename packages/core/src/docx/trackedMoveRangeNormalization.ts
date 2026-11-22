import { panic } from "better-result";

import { PARSE_WARNING_CODES } from "@stll/docx-core/model";

import type {
  BlockContent,
  DocumentBody,
  Endnote,
  Footnote,
  HeaderFooter,
} from "../types/document";
import { rebuildBodySections } from "./documentParser";
import { InlineContentRemovals } from "./inlineContentRemovals";
import {
  type DocxContentSlot,
  visitDocxContentSlots,
} from "./paragraphTraversal";

/** The code this normalisation is reported under, owned here, not at the caller. */
export const UNBALANCED_MOVE_RANGE_WARNING: "unbalanced-move-range" =
  PARSE_WARNING_CODES.unbalancedMoveRange;

type NormalizeTrackedMoveRangesInput = {
  documentBody: DocumentBody;
  headers?: Map<string, HeaderFooter>;
  footers?: Map<string, HeaderFooter>;
  footnotes?: readonly Footnote[];
  endnotes?: readonly Endnote[];
};

/** The code a dropped further range is reported under. */
export const DROPPED_MOVE_RANGE_WARNING: "dropped-move-range" =
  PARSE_WARNING_CODES.droppedMoveRange;

/** A move range after the first of its kind and id, dropped. */
type DroppedMoveRange = { kind: "moveFrom" | "moveTo"; id: number };

type NormalizeTrackedMoveRangesResult = {
  removedUnbalancedMoveRangeMarkers: number;
  /** Full ranges after a move's first, dropped: a move holds one range. */
  droppedMoveRanges?: DroppedMoveRange[];
};

type MoveRangeMarkerRef = {
  slot: DocxContentSlot;
  id: number;
  kind: "moveFrom" | "moveTo";
  side: "start" | "end";
};

export const normalizeTrackedMoveRanges = ({
  documentBody,
  headers,
  footers,
  footnotes,
  endnotes,
}: NormalizeTrackedMoveRangesInput): NormalizeTrackedMoveRangesResult => {
  const rangeMarkers: MoveRangeMarkerRef[] = [];

  visitDocxContentSlots(
    { documentBody, headers, footers, footnotes, endnotes },
    (slot) => {
      const marker = toMoveRangeMarkerRef(slot);
      if (marker) {
        rangeMarkers.push(marker);
      }
    },
  );
  // A header, footer or note numbers its ranges apart from the body.
  const partOf = new Map<unknown, unknown>();
  if (rangeMarkers.length > 0)
    for (const surfaces of [
      ...[...(headers?.values() ?? [])].map((part) => ({
        headers: new Map([["part", part]]),
      })),
      ...[...(footers?.values() ?? [])].map((part) => ({
        footers: new Map([["part", part]]),
      })),
      ...(footnotes ?? []).map((note) => ({ footnotes: [note] })),
      ...(endnotes ?? []).map((note) => ({ endnotes: [note] })),
    ])
      visitDocxContentSlots(
        { documentBody: { content: [] }, ...surfaces },
        ({ content }) => partOf.set(content, surfaces),
      );

  return removeUnbalancedMoveRangeMarkers({
    rangeMarkers,
    documentBody,
    partOf: ({ slot }) => partOf.get(slot.content) ?? documentBody,
  });
};

type RemoveUnbalancedMoveRangeMarkersOptions = {
  rangeMarkers: readonly MoveRangeMarkerRef[];
  documentBody: DocumentBody;
  /** The part a marker stands in: a range repeats one of its own part. */
  partOf: (marker: MoveRangeMarkerRef) => unknown;
};

const removeUnbalancedMoveRangeMarkers = ({
  rangeMarkers,
  documentBody,
  partOf,
}: RemoveUnbalancedMoveRangeMarkersOptions): NormalizeTrackedMoveRangesResult => {
  const byRange = new Map<string, MoveRangeMarkerRef[]>();
  for (const marker of rangeMarkers) {
    const key = `${marker.kind}:${marker.id}`;
    const markers = byRange.get(key);
    if (markers) {
      markers.push(marker);
      continue;
    }
    byRange.set(key, [marker]);
  }

  const removals = new InlineContentRemovals();
  const blockRemovals = new Map<BlockContent[], Set<number>>();
  const markForRemoval = ({ slot }: MoveRangeMarkerRef): void => {
    switch (slot.type) {
      case "inline":
        removals.mark(slot);
        return;
      case "block": {
        const indexes = blockRemovals.get(slot.content);
        if (indexes) {
          indexes.add(slot.index);
          return;
        }
        blockRemovals.set(slot.content, new Set([slot.index]));
        return;
      }
      default: {
        const unsupported: never = slot;
        panic(
          `Unsupported move range ownership: ${JSON.stringify(unsupported)}`,
        );
      }
    }
  };

  const droppedMoveRanges: DroppedMoveRange[] = [];
  let droppedMarkers = 0;
  for (const markers of byRange.values()) {
    let openStart: MoveRangeMarkerRef | null = null;
    const held = new Set<unknown>();
    for (const marker of markers) {
      if (marker.side === "start") {
        if (openStart) {
          markForRemoval(marker);
          continue;
        }
        openStart = marker;
        continue;
      }

      if (!openStart) {
        markForRemoval(marker);
        continue;
      }
      // A part's first full range is the move's; one after it repeats it.
      const part = partOf(marker);
      if (held.has(part)) {
        markForRemoval(openStart);
        markForRemoval(marker);
        droppedMarkers += 2;
        droppedMoveRanges.push({ kind: marker.kind, id: marker.id });
      }
      held.add(part);
      openStart = null;
    }

    if (openStart) {
      markForRemoval(openStart);
    }
  }

  let removedMarkers = removals.apply();
  for (const [content, indexes] of blockRemovals) {
    const kept = content.filter((_, index) => !indexes.has(index));
    removedMarkers += content.length - kept.length;
    content.length = 0;
    content.push(...kept);
  }

  if (blockRemovals.has(documentBody.content))
    rebuildBodySections(documentBody);

  return {
    removedUnbalancedMoveRangeMarkers: removedMarkers - droppedMarkers,
    ...(droppedMoveRanges.length === 0 ? {} : { droppedMoveRanges }),
  };
};

const toMoveRangeMarkerRef = (
  slot: DocxContentSlot,
): MoveRangeMarkerRef | null => {
  const marker = slot.item;
  if (marker.type === "moveFromRangeStart") {
    return {
      slot,
      id: marker.id,
      kind: "moveFrom",
      side: "start",
    };
  }
  if (marker.type === "moveFromRangeEnd") {
    return {
      slot,
      id: marker.id,
      kind: "moveFrom",
      side: "end",
    };
  }
  if (marker.type === "moveToRangeStart") {
    return {
      slot,
      id: marker.id,
      kind: "moveTo",
      side: "start",
    };
  }
  if (marker.type === "moveToRangeEnd") {
    return {
      slot,
      id: marker.id,
      kind: "moveTo",
      side: "end",
    };
  }
  return null;
};
