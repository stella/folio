import type { ListRendering, ParagraphFormatting } from "../types/document";
import type { ParagraphAttrs } from "./schema/nodes";

type ListRenderingNumPrField = "level" | "numId";
type ListRenderingPresentationField = Exclude<keyof ListRendering, ListRenderingNumPrField>;

type CompatibleParagraphAttr<Field extends ListRenderingPresentationField> = {
  [Attr in keyof ParagraphAttrs]: Exclude<
    ParagraphAttrs[Attr],
    null | undefined
  > extends NonNullable<ListRendering[Field]>
    ? NonNullable<ListRendering[Field]> extends Exclude<ParagraphAttrs[Attr], null | undefined>
      ? Attr
      : never
    : never;
}[keyof ParagraphAttrs];

type ListRenderingFieldDisposition = {
  [Field in ListRenderingNumPrField]: { readonly type: "numPr" };
} & {
  [Field in ListRenderingPresentationField]: {
    readonly type: "attr";
    readonly attr: CompatibleParagraphAttr<Field>;
    readonly presence: "defined" | "truthy";
  };
};

/**
 * Total ownership map for the list presentation transported through paragraph
 * attrs. Extending ListRendering requires assigning the new field to numPr or
 * a type-compatible attr before the model compiles.
 */
const LIST_RENDERING_FIELD_DISPOSITIONS = {
  marker: { type: "attr", attr: "listMarker", presence: "truthy" },
  markerTemplate: { type: "attr", attr: "listMarkerTemplate", presence: "truthy" },
  level: { type: "numPr" },
  numId: { type: "numPr" },
  isBullet: { type: "attr", attr: "listIsBullet", presence: "truthy" },
  isLegal: { type: "attr", attr: "listIsLegal", presence: "truthy" },
  numFmt: { type: "attr", attr: "listNumFmt", presence: "truthy" },
  markerHidden: { type: "attr", attr: "listMarkerHidden", presence: "truthy" },
  markerFormatting: {
    type: "attr",
    attr: "listMarkerFormatting",
    presence: "truthy",
  },
  markerAlignment: { type: "attr", attr: "listMarkerAlignment", presence: "truthy" },
  markerAllCaps: { type: "attr", attr: "listMarkerAllCaps", presence: "truthy" },
  markerSuffix: { type: "attr", attr: "listMarkerSuffix", presence: "truthy" },
  levelNumFmts: { type: "attr", attr: "listLevelNumFmts", presence: "truthy" },
  levelStarts: { type: "attr", attr: "listLevelStarts", presence: "defined" },
  abstractNumId: { type: "attr", attr: "listAbstractNumId", presence: "defined" },
  startOverride: { type: "attr", attr: "listStartOverride", presence: "defined" },
  implicitChildLevelAdvances: {
    type: "attr",
    attr: "listImplicitChildLevelAdvances",
    presence: "defined",
  },
  markerSecondSlotOffsetTwips: {
    type: "attr",
    attr: "listMarkerSecondSlotOffsetTwips",
    presence: "defined",
  },
} as const satisfies ListRenderingFieldDisposition;

/** Derive the canonical PM attr patch from the total list-rendering ownership map. */
export const listRenderingAttrPatch = (
  rendering: ListRendering | undefined,
): Partial<ParagraphAttrs> => {
  const attrs: Partial<ParagraphAttrs> = {};
  if (!rendering) return attrs;

  for (const [field, disposition] of Object.entries(LIST_RENDERING_FIELD_DISPOSITIONS)) {
    if (disposition.type === "numPr") continue;
    const value = Reflect.get(rendering, field);
    if (disposition.presence === "truthy" ? !value : value === undefined) continue;
    Reflect.set(attrs, disposition.attr, value);
  }
  return attrs;
};

/** Replace the direct `w:numPr` owned by the paragraph serializer. */
export const withDirectListNumbering = (
  formatting: ParagraphFormatting | null | undefined,
  numPr: ParagraphFormatting["numPr"] | null | undefined,
): ParagraphFormatting | undefined => {
  const result = { ...formatting };
  if (numPr === null || numPr === undefined) {
    Reflect.deleteProperty(result, "numPr");
  } else {
    result.numPr = { ...numPr };
  }
  return Object.keys(result).length > 0 ? result : undefined;
};

/** Project every modeled list-presentation field into its canonical PM attr. */
export const applyListRenderingAttrs = (
  attrs: ParagraphAttrs,
  rendering: ListRendering | undefined,
): void => {
  Object.assign(attrs, listRenderingAttrPatch(rendering));
};

/** Restore every modeled list-presentation field from its canonical PM attr. */
export const listRenderingFromAttrs = (attrs: ParagraphAttrs): ListRendering | undefined => {
  const numId = attrs.numPr?.numId;
  if (numId === undefined || numId === 0) return undefined;

  const hasRenderingInfo =
    attrs.listMarker != null || attrs.listIsBullet || attrs.listNumFmt != null;
  if (!hasRenderingInfo) return undefined;

  const rendering: ListRendering = {
    marker: attrs.listMarker ?? "",
    level: attrs.numPr?.ilvl ?? 0,
    numId,
    isBullet: attrs.listIsBullet ?? false,
  };
  for (const [field, disposition] of Object.entries(LIST_RENDERING_FIELD_DISPOSITIONS)) {
    if (disposition.type === "numPr" || field === "marker" || field === "isBullet") continue;
    const value = Reflect.get(attrs, disposition.attr);
    if (value == null) continue;
    Reflect.set(rendering, field, value);
  }
  return rendering;
};
