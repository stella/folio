import { PARSE_WARNING_CODES } from "@stll/docx-core/model";

import type {
  DocumentBody,
  Endnote,
  Footnote,
  HeaderFooter,
  ParagraphContent,
} from "../types/document";
import {
  InlineContentRemovals,
  visitDocxParagraphs,
  visitInlineContentSlots,
} from "./paragraphTraversal";

/** The code this normalisation is reported under, owned here, not at the caller. */
export const UNBALANCED_MOVE_RANGE_WARNING = PARSE_WARNING_CODES.unbalancedMoveRange;

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
  content: ParagraphContent[];
  index: number;
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

  visitDocxParagraphs({ documentBody, headers, footers, footnotes, endnotes }, (paragraph) => {
    visitInlineContentSlots(paragraph, ({ content, index, item }) => {
      const marker = toMoveRangeMarkerRef(content, index, item);
      if (marker) {
        rangeMarkers.push(marker);
      }
    });
  });

  return {
    removedUnbalancedMoveRangeMarkers: removeUnbalancedMoveRangeMarkers(rangeMarkers),
  };
};

const removeUnbalancedMoveRangeMarkers = (rangeMarkers: readonly MoveRangeMarkerRef[]): number => {
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
  const markForRemoval = (marker: MoveRangeMarkerRef): void => {
    removals.mark(marker);
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

  return removals.apply();
};

const toMoveRangeMarkerRef = (
  content: ParagraphContent[],
  index: number,
  marker: ParagraphContent,
): MoveRangeMarkerRef | null => {
  if (marker.type === "moveFromRangeStart") {
    return {
      content,
      index,
      id: marker.id,
      kind: "moveFrom",
      side: "start",
    };
  }
  if (marker.type === "moveFromRangeEnd") {
    return {
      content,
      index,
      id: marker.id,
      kind: "moveFrom",
      side: "end",
    };
  }
  if (marker.type === "moveToRangeStart") {
    return {
      content,
      index,
      id: marker.id,
      kind: "moveTo",
      side: "start",
    };
  }
  if (marker.type === "moveToRangeEnd") {
    return {
      content,
      index,
      id: marker.id,
      kind: "moveTo",
      side: "end",
    };
  }
  return null;
};
