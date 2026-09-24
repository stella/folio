/**
 * The one table that ties `Paragraph.listRendering` to the editor's `list*`
 * attrs. Both projections (model → attrs in `toProseDoc`, attrs → model in
 * `fromProseDoc`) and the cleared-attr set derive from it, so a field added
 * to `ListRendering` without a projection fails to compile instead of being
 * silently dropped on the way out of the editor.
 *
 * `level` and `numId` are not here: they travel in `numPr`.
 */

import { paragraphNumberingLevel } from "@stll/docx-core/model";

import type { ListRendering } from "../types/document";
import type { ParagraphAttrs } from "./schema/nodes";

/** `ListRendering` fields that have their own editor attr. */
export type ListRenderingAttrField = Exclude<keyof ListRendering, "level" | "numId">;

export const LIST_RENDERING_ATTR_BY_FIELD = {
  marker: "listMarker",
  markerTemplate: "listMarkerTemplate",
  isBullet: "listIsBullet",
  isLegal: "listIsLegal",
  numFmt: "listNumFmt",
  markerHidden: "listMarkerHidden",
  markerFormatting: "listMarkerFormatting",
  markerAlignment: "listMarkerAlignment",
  markerAllCaps: "listMarkerAllCaps",
  markerSuffix: "listMarkerSuffix",
  levelTabs: "listLevelTabs",
  levelNumFmts: "listLevelNumFmts",
  levelStarts: "listLevelStarts",
  abstractNumId: "listAbstractNumId",
  startOverride: "listStartOverride",
  implicitChildLevelAdvances: "listImplicitChildLevelAdvances",
  markerSecondSlotOffsetTwips: "listMarkerSecondSlotOffsetTwips",
} as const satisfies Record<ListRenderingAttrField, keyof ParagraphAttrs>;

export type ListRenderingAttrKey = (typeof LIST_RENDERING_ATTR_BY_FIELD)[ListRenderingAttrField];

export const LIST_RENDERING_ATTR_KEYS: readonly ListRenderingAttrKey[] = Object.values(
  LIST_RENDERING_ATTR_BY_FIELD,
);

/** Every list attr, required, so a table entry without a projection is a compile error. */
type ListRenderingAttrPatch = { [K in ListRenderingAttrKey]-?: ParagraphAttrs[K] | undefined };

type OptionalListRenderingField = Exclude<ListRenderingAttrField, "marker" | "isBullet">;

/** Every optional `ListRendering` field, required, for the same reason. */
type OptionalListRenderingFields = {
  [F in OptionalListRenderingField]-?: ListRendering[F] | undefined;
};

const isDefinedEntry = ([, value]: [string, unknown]): boolean => value !== undefined;

const definedAttrs = (
  patch: ListRenderingAttrPatch,
): Partial<Pick<ParagraphAttrs, ListRenderingAttrKey>> =>
  Object.fromEntries(Object.entries(patch).filter(isDefinedEntry));

const definedFields = (
  fields: OptionalListRenderingFields,
): Partial<Pick<ListRendering, OptionalListRenderingField>> =>
  Object.fromEntries(Object.entries(fields).filter(isDefinedEntry));

/**
 * Editor attrs carrying `rendering`. A falsy marker, flag or format is left
 * unset: the editor treats `null` and absent alike for those, and the parser
 * never authors them as `false` or `""` deliberately.
 */
export const listRenderingAttrPatch = (rendering: ListRendering): Partial<ParagraphAttrs> => {
  const patch: ListRenderingAttrPatch = {
    listMarker: rendering.marker || undefined,
    listMarkerTemplate: rendering.markerTemplate || undefined,
    listIsBullet: rendering.isBullet || undefined,
    listIsLegal: rendering.isLegal || undefined,
    listNumFmt: rendering.numFmt || undefined,
    listMarkerHidden: rendering.markerHidden || undefined,
    listMarkerFormatting: rendering.markerFormatting || undefined,
    listMarkerAlignment: rendering.markerAlignment || undefined,
    listMarkerAllCaps: rendering.markerAllCaps || undefined,
    listMarkerSuffix: rendering.markerSuffix || undefined,
    listLevelTabs: rendering.levelTabs?.length ? rendering.levelTabs : undefined,
    listLevelNumFmts: rendering.levelNumFmts || undefined,
    listLevelStarts: rendering.levelStarts || undefined,
    listAbstractNumId: rendering.abstractNumId,
    listStartOverride: rendering.startOverride,
    listImplicitChildLevelAdvances: rendering.implicitChildLevelAdvances,
    listMarkerSecondSlotOffsetTwips: rendering.markerSecondSlotOffsetTwips,
  };
  return definedAttrs(patch);
};

type ListRenderingFromAttrsOptions = {
  attrs: ParagraphAttrs;
  numId: number;
};

/**
 * Inverse of `listRenderingAttrPatch`. The caller decides whether the
 * paragraph is a list item at all; this only reassembles the fields.
 */
export const listRenderingFromAttrs = ({
  attrs,
  numId,
}: ListRenderingFromAttrsOptions): ListRendering => {
  const optional: OptionalListRenderingFields = {
    markerTemplate: attrs.listMarkerTemplate ?? undefined,
    isLegal: attrs.listIsLegal ?? undefined,
    numFmt: attrs.listNumFmt ?? undefined,
    markerHidden: attrs.listMarkerHidden ?? undefined,
    markerFormatting: attrs.listMarkerFormatting ?? undefined,
    markerAlignment: attrs.listMarkerAlignment ?? undefined,
    markerAllCaps: attrs.listMarkerAllCaps ?? undefined,
    markerSuffix: attrs.listMarkerSuffix ?? undefined,
    levelTabs: attrs.listLevelTabs ?? undefined,
    levelNumFmts: attrs.listLevelNumFmts ?? undefined,
    levelStarts: attrs.listLevelStarts ?? undefined,
    abstractNumId: attrs.listAbstractNumId ?? undefined,
    startOverride: attrs.listStartOverride ?? undefined,
    implicitChildLevelAdvances: attrs.listImplicitChildLevelAdvances ?? undefined,
    markerSecondSlotOffsetTwips: attrs.listMarkerSecondSlotOffsetTwips ?? undefined,
  };
  return {
    marker: attrs.listMarker ?? "",
    level: paragraphNumberingLevel(attrs.numPr) ?? 0,
    numId,
    isBullet: attrs.listIsBullet ?? false,
    ...definedFields(optional),
  };
};
