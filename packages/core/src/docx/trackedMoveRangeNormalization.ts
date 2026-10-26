import { panic } from "better-result";

import { PARSE_WARNING_CODES } from "@stll/docx-core/model";

import type {
  BlockContent,
  DocumentBody,
  Endnote,
  Footnote,
  HeaderFooter,
} from "../types/document";
import { buildSections } from "./documentParser";
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

type NormalizeTrackedMoveRangesResult = {
  removedUnbalancedMoveRangeMarkers: number;
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

  return removeUnbalancedMoveRangeMarkers({ rangeMarkers, documentBody });
};

type RemoveUnbalancedMoveRangeMarkersOptions = {
  rangeMarkers: readonly MoveRangeMarkerRef[];
  documentBody: DocumentBody;
};

const removeUnbalancedMoveRangeMarkers = ({
  rangeMarkers,
  documentBody,
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

  for (const markers of byRange.values()) {
    let openStart: MoveRangeMarkerRef | null = null;
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
      openStart = null;
    }

    if (openStart) {
      markForRemoval(openStart);
    }
  }

  let removedUnbalancedMoveRangeMarkers = removals.apply();
  for (const [content, indexes] of blockRemovals) {
    const kept = content.filter((_, index) => !indexes.has(index));
    removedUnbalancedMoveRangeMarkers += content.length - kept.length;
    content.length = 0;
    content.push(...kept);
  }

  if (blockRemovals.has(documentBody.content) && documentBody.sections) {
    const rebuiltSections = buildSections(
      documentBody.content,
      documentBody.finalSectionProperties,
    );
    if (rebuiltSections.length !== documentBody.sections.length) {
      panic("Removing move range markers changed the document section count");
    }
    for (const [index, section] of documentBody.sections.entries()) {
      const rebuilt = rebuiltSections.at(index);
      if (!rebuilt) {
        panic("Missing rebuilt section after move range normalization");
      }
      section.content = rebuilt.content;
    }
  }

  return { removedUnbalancedMoveRangeMarkers };
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
