import type { ListLevel, NumberingDefinitions } from "./lists";

const MIN_NUMBERING_LEVEL = 0;
const MAX_NUMBERING_LEVEL = 8;

export type NumberingLevelResolution =
  | { kind: "level"; level: ListLevel; source: "abstract" | "override" | "styleLink" }
  | {
      kind: "undefined";
      reason:
        | "levelOutOfRange"
        | "noInstance"
        | "noAbstract"
        | "unresolvedStyleLink"
        | "levelNotDefined";
    };

export type NumberingLevelEntry = { numId: number; ilvl: number; level: ListLevel };

/** Whether a numbering level reserves space for a visible marker. */
export const numberingLevelHasMarkerSlot = (level: Pick<ListLevel, "numFmt">): boolean =>
  level.numFmt !== "none";

const firstById = <Value>(values: readonly Value[], idOf: (value: Value) => number, id: number) =>
  values.find((value) => idOf(value) === id);

/**
 * Resolves the level named by a concrete numbering reference.
 *
 * ECMA-376 §17.9.3 and §17.9.6 restrict `ilvl` to 0..8; an invalid or
 * undefined level remains authored but is not materialised. Sections
 * §17.9.8 and §17.9.9 make a full `lvlOverride/lvl` authoritative while a
 * `startOverride` changes the selected full or abstract level's start. Sections
 * §17.9.21 and §17.9.27 link a level-less abstract through `numStyleLink` to
 * the first document-order abstract whose `styleLink` matches.
 *
 * Duplicate ids are resolved in document order, matching the package part.
 */
export const resolveNumberingLevel = (
  definitions: NumberingDefinitions,
  numId: number,
  ilvl: number,
): NumberingLevelResolution => {
  if (!Number.isInteger(ilvl) || ilvl < MIN_NUMBERING_LEVEL || ilvl > MAX_NUMBERING_LEVEL) {
    return { kind: "undefined", reason: "levelOutOfRange" };
  }
  const instance = firstById(definitions.nums, (candidate) => candidate.numId, numId);
  if (!instance) return { kind: "undefined", reason: "noInstance" };

  // A full level override is a complete level definition of its own, so it
  // resolves even when the instance's abstract definition is missing.
  const override = instance.levelOverrides?.find((candidate) => candidate.ilvl === ilvl);
  if (override?.lvl) {
    return {
      kind: "level",
      level:
        override.startOverride === undefined
          ? override.lvl
          : { ...override.lvl, start: override.startOverride },
      source: "override",
    };
  }

  const abstract = firstById(
    definitions.abstractNums,
    (candidate) => candidate.abstractNumId,
    instance.abstractNumId,
  );
  if (!abstract) return { kind: "undefined", reason: "noAbstract" };

  let levelOwner = abstract;
  let source: "abstract" | "styleLink" = "abstract";
  if (abstract.levels.length === 0 && abstract.numStyleLink !== undefined) {
    const linked = definitions.abstractNums.find(
      (candidate) => candidate.styleLink === abstract.numStyleLink,
    );
    if (!linked) return { kind: "undefined", reason: "unresolvedStyleLink" };
    levelOwner = linked;
    source = "styleLink";
  }

  const level = levelOwner.levels.find((candidate) => candidate.ilvl === ilvl);
  if (!level) return { kind: "undefined", reason: "levelNotDefined" };
  if (override?.startOverride === undefined) return { kind: "level", level, source };
  return {
    kind: "level",
    level: { ...level, start: override.startOverride },
    source,
  };
};

/** ECMA-376 §17.9.6 permits nine levels; enumerate each instance through the owner. */
export const numberingLevelEntries = (definitions: NumberingDefinitions): NumberingLevelEntry[] => {
  const entries: NumberingLevelEntry[] = [];
  const seen = new Set<number>();
  for (const instance of definitions.nums) {
    if (seen.has(instance.numId)) continue;
    seen.add(instance.numId);
    for (let ilvl = MIN_NUMBERING_LEVEL; ilvl <= MAX_NUMBERING_LEVEL; ilvl += 1) {
      const resolution = resolveNumberingLevel(definitions, instance.numId, ilvl);
      if (resolution.kind === "level") {
        entries.push({ numId: instance.numId, ilvl, level: resolution.level });
      }
    }
  }
  return entries;
};

/**
 * Whether the resolved level paints a non-blank marker. ECMA-376 §17.3.2.41
 * hides vanished runs; §17.9.11, §17.9.14 and §17.9.23 define the
 * current level's legal-numbering mode, `none`, and placeholder rendering.
 */
export const levelShowsMarker = (
  definitions: NumberingDefinitions,
  numId: number,
  ilvl: number,
): boolean => {
  const resolution = resolveNumberingLevel(definitions, numId, ilvl);
  if (resolution.kind === "undefined" || resolution.level.rPr?.hidden === true) return false;
  const { level } = resolution;
  if (level.lvlText.length === 0) return level.isLgl === true || numberingLevelHasMarkerSlot(level);

  const rendered = level.lvlText.replace(/%([1-9])/gu, (_placeholder, ordinal: string) => {
    const referenced = resolveNumberingLevel(definitions, numId, Number.parseInt(ordinal, 10) - 1);
    if (referenced.kind === "undefined") return "";
    return level.isLgl === true || numberingLevelHasMarkerSlot(referenced.level) ? "x" : "";
  });
  return rendered.trim().length > 0;
};
