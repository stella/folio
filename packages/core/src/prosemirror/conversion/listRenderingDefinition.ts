import type { ListRendering } from "../../types/document";
import { canonicalJson } from "../../utils/canonicalJson";

// Paragraph counters and folded fields do not belong to the numbering definition.
const RENDERING_FIELDS = {
  marker: { key: "marker", source: "paragraph" },
  implicitChildLevelAdvances: { key: "implicitChildLevelAdvances", source: "paragraph" },
  markerSecondSlotOffsetTwips: { key: "markerSecondSlotOffsetTwips", source: "paragraph" },
  markerTemplate: { key: "markerTemplate", source: "definition" },
  level: { key: "level", source: "definition" },
  numId: { key: "numId", source: "definition" },
  isBullet: { key: "isBullet", source: "definition" },
  isLegal: { key: "isLegal", source: "definition" },
  numFmt: { key: "numFmt", source: "definition" },
  markerHidden: { key: "markerHidden", source: "definition" },
  markerFormatting: { key: "markerFormatting", source: "definition" },
  markerAlignment: { key: "markerAlignment", source: "definition" },
  markerAllCaps: { key: "markerAllCaps", source: "definition" },
  markerSuffix: { key: "markerSuffix", source: "definition" },
  levelTabs: { key: "levelTabs", source: "definition" },
  levelNumFmts: { key: "levelNumFmts", source: "definition" },
  levelStarts: { key: "levelStarts", source: "definition" },
  abstractNumId: { key: "abstractNumId", source: "definition" },
  startOverride: { key: "startOverride", source: "definition" },
} as const satisfies {
  [Key in keyof ListRendering]-?: { key: Key; source: "paragraph" | "definition" };
};

const DEFINITION_FIELDS = Object.values(RENDERING_FIELDS).filter(
  (field) => field.source === "definition",
);

/** Compare scalar fields directly; only nested records need canonical ordering. */
export const listRenderingDefinitionsMatch = (
  left: ListRendering,
  right: ListRendering,
): boolean => {
  for (const { key } of DEFINITION_FIELDS) {
    const before = left[key];
    const after = right[key];
    if (before === after) continue;
    if (typeof before !== "object" || typeof after !== "object") return false;
    if (canonicalJson(before) !== canonicalJson(after)) return false;
  }
  return true;
};
