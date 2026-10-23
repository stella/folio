/**
 * List-marker resolution for paragraphs, including tracked numbering changes
 * (inserted, removed, and changed list markers).
 */

import { sameStatedParagraphNumbering } from "@stll/docx-core/model";

import {
  bulletMarkerFontName,
  convertBulletToUnicode,
} from "../../docx/bulletMarkers";
import type { FontAlternates } from "../../fonts/fontAlternates";
import type { ParagraphAttrs } from "../../layout-engine/types";
import { isListNumPr } from "../../layout-engine/types";
import type {
  ParagraphAttrs as PMParagraphAttrs,
  ParagraphPropertyChangeAttrs,
} from "../../prosemirror/schema/nodes";
import type { TextFormatting, Theme } from "../../types/document";
import { listMarkerFormattingFor } from "./textFormattingConversion";

export function applyMarkerAllCaps(
  marker: string | null,
  allCaps: boolean | undefined,
): string | null {
  if (marker === null || !allCaps) {
    return marker;
  }
  return marker.toLocaleUpperCase();
}

/**
 * Convert PM paragraph attrs to layout engine paragraph attrs.
 */
type ListMarkerRevision = NonNullable<ParagraphAttrs["listMarkerRevision"]>;

export type ListPropertyChange = ParagraphPropertyChangeAttrs;
export type ListPropertyFormatting = NonNullable<
  ListPropertyChange["previousFormatting"]
>;

export function toListMarkerRevision(
  kind: ListMarkerRevision["kind"],
  info: ListPropertyChange["info"],
): ListMarkerRevision {
  const revision: ListMarkerRevision = { kind };
  if (typeof info?.author === "string") {
    revision.author = info.author;
  }
  if (typeof info?.date === "string") {
    revision.date = info.date;
  }
  if (typeof info?.id === "number") {
    revision.revisionId = info.id;
  }
  return revision;
}

export function isAddedNumberingChange(
  change: ListPropertyChange,
): change is ListPropertyChange & {
  previousFormatting: ListPropertyFormatting;
} {
  const previousFormatting = change.previousFormatting;
  return (
    previousFormatting != null &&
    Object.hasOwn(previousFormatting, "numPr") &&
    previousFormatting.numPr == null
  );
}

export function isRemovedNumberingChange(
  change: ListPropertyChange,
): change is ListPropertyChange & {
  previousFormatting: ListPropertyFormatting;
} {
  const previousFormatting = change.previousFormatting;
  return (
    previousFormatting != null &&
    Object.hasOwn(previousFormatting, "numPr") &&
    isListNumPr(previousFormatting.numPr)
  );
}

export function isChangedNumberingChange(
  currentNumPr: NonNullable<PMParagraphAttrs["numPr"]>,
  change: ListPropertyChange,
): change is ListPropertyChange & {
  previousFormatting: ListPropertyFormatting;
} {
  const previousFormatting = change.previousFormatting;
  return (
    previousFormatting != null &&
    Object.hasOwn(previousFormatting, "numPr") &&
    isListNumPr(previousFormatting.numPr) &&
    !sameStatedParagraphNumbering(previousFormatting.numPr, currentNumPr)
  );
}

function toPreviousListAttrs(
  previousFormatting: ListPropertyFormatting,
): PMParagraphAttrs {
  const attrs: PMParagraphAttrs = {};
  const numPr = previousFormatting.numPr;
  if (isListNumPr(numPr)) {
    attrs.numPr = numPr;
  }

  const listIsBullet = previousFormatting.listIsBullet;
  if (listIsBullet !== undefined) {
    attrs.listIsBullet = listIsBullet;
  }

  const listIsLegal = previousFormatting.listIsLegal;
  if (listIsLegal !== undefined) {
    attrs.listIsLegal = listIsLegal;
  }

  const listMarker = previousFormatting.listMarker;
  if (listMarker !== undefined) {
    attrs.listMarker = listMarker;
  }

  const listMarkerTemplate = previousFormatting.listMarkerTemplate;
  if (listMarkerTemplate !== undefined) {
    attrs.listMarkerTemplate = listMarkerTemplate;
  }

  const listNumFmt = previousFormatting.listNumFmt;
  if (listNumFmt !== undefined) {
    attrs.listNumFmt = listNumFmt;
  }

  const listLevelNumFmts = previousFormatting.listLevelNumFmts;
  if (listLevelNumFmts !== undefined) {
    attrs.listLevelNumFmts = listLevelNumFmts;
  }

  const listLevelStarts = previousFormatting.listLevelStarts;
  if (listLevelStarts !== undefined) {
    attrs.listLevelStarts = listLevelStarts;
  }

  const listAbstractNumId = previousFormatting.listAbstractNumId;
  if (listAbstractNumId !== undefined) {
    attrs.listAbstractNumId = listAbstractNumId;
  }

  const listStartOverride = previousFormatting.listStartOverride;
  if (listStartOverride !== undefined) {
    attrs.listStartOverride = listStartOverride;
  }

  const listMarkerHidden = previousFormatting.listMarkerHidden;
  if (listMarkerHidden !== undefined) {
    attrs.listMarkerHidden = listMarkerHidden;
  }

  const listMarkerFormatting = previousFormatting.listMarkerFormatting;
  if (listMarkerFormatting !== undefined) {
    attrs.listMarkerFormatting = listMarkerFormatting;
  }

  const listMarkerAlignment = previousFormatting.listMarkerAlignment;
  if (listMarkerAlignment !== undefined) {
    attrs.listMarkerAlignment = listMarkerAlignment;
  }

  const listMarkerSuffix = previousFormatting.listMarkerSuffix;
  if (listMarkerSuffix !== undefined) {
    attrs.listMarkerSuffix = listMarkerSuffix;
  }

  const listMarkerAllCaps = previousFormatting.listMarkerAllCaps;
  if (listMarkerAllCaps !== undefined) {
    attrs.listMarkerAllCaps = listMarkerAllCaps;
  }

  const listImplicitChildLevelAdvances =
    previousFormatting.listImplicitChildLevelAdvances;
  if (listImplicitChildLevelAdvances !== undefined) {
    attrs.listImplicitChildLevelAdvances = listImplicitChildLevelAdvances;
  }

  const listMarkerSecondSlotOffsetTwips =
    previousFormatting.listMarkerSecondSlotOffsetTwips;
  if (listMarkerSecondSlotOffsetTwips !== undefined) {
    attrs.listMarkerSecondSlotOffsetTwips = listMarkerSecondSlotOffsetTwips;
  }

  return attrs;
}

function resolveDeletedListMarker(
  previousListAttrs: PMParagraphAttrs,
): string | null {
  if (previousListAttrs.listMarker) {
    return previousListAttrs.listIsBullet
      ? convertBulletToUnicode(
          previousListAttrs.listMarker,
          bulletMarkerFontName(previousListAttrs.listMarkerFormatting),
        )
      : previousListAttrs.listMarker;
  }

  if (previousListAttrs.listIsBullet) {
    return "\u2022";
  }

  return null;
}

export function applyDeletedListMarkerAttrs(
  attrs: ParagraphAttrs,
  change: ListPropertyChange & { previousFormatting: ListPropertyFormatting },
  theme: Theme | null | undefined,
  fontAlternates: FontAlternates | undefined,
  paragraphMarkFormatting: () => TextFormatting | undefined,
): void {
  const previousListAttrs = toPreviousListAttrs(change.previousFormatting);
  const marker = resolveDeletedListMarker(previousListAttrs);
  if (!marker) {
    return;
  }

  attrs.listMarker = marker;
  attrs.listMarkerRevision = toListMarkerRevision("del", change.info);
  if (previousListAttrs.listIsBullet !== undefined) {
    attrs.listIsBullet = previousListAttrs.listIsBullet;
  }
  if (previousListAttrs.listMarkerHidden !== undefined) {
    attrs.listMarkerHidden = previousListAttrs.listMarkerHidden;
  }
  const listMarkerFormatting = listMarkerFormattingFor(
    previousListAttrs.listMarkerFormatting,
    paragraphMarkFormatting(),
    theme,
    fontAlternates,
  );
  if (listMarkerFormatting) {
    attrs.listMarkerFormatting = listMarkerFormatting;
  }
  if (previousListAttrs.listMarkerAlignment) {
    attrs.listMarkerAlignment = previousListAttrs.listMarkerAlignment;
  }
  if (previousListAttrs.listMarkerSuffix) {
    attrs.listMarkerSuffix = previousListAttrs.listMarkerSuffix;
  }
}
