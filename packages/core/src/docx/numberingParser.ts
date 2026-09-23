/**
 * Numbering/List Parser for DOCX
 *
 * Parses numbering.xml to extract:
 * - Abstract numbering definitions (templates with levels)
 * - Numbering instances (concrete references with optional overrides)
 *
 * OOXML Structure:
 * - w:abstractNum - Template definitions with 9 levels (0-8)
 * - w:num - Instances that reference abstractNum and can override levels
 * - w:lvl - Level definition with start, format, text pattern, etc.
 */

import type {
  NumberingDefinitions,
  AbstractNumbering,
  LevelLegacy,
  LevelOverride,
  NumberingInstance,
  ListLevel,
  ListRendering,
  CounterFormat,
  ParagraphAlignment,
  TextFormatting,
  ListMarkerFormatting,
} from "../types/document";
import { attributeRemainder, DERIVED_PART_ROOT_ATTRIBUTES } from "./attributeRemainder";
import {
  CAPTURE,
  type ChildHandlers,
  type ChildReader,
  dispatchChildrenWithContext,
  sequencePositions,
} from "./containerChildren";
import { isNumberingReference } from "./numberingReference";
import { formatOoxmlCounter } from "./ooxmlCounterFormatter";
import { parseParagraphProperties } from "./paragraphProperties";
import {
  LevelSuffixSchema,
  narrowEnum,
  NumberFormatSchema,
  ParagraphAlignmentSchema,
} from "./parserEnums";
import { parseRunProperties, RUN_PROPERTY_OWNERS } from "./runParser";
import {
  parseXmlDocument,
  findChild,
  getAttribute,
  parseBooleanElement,
  parseNumericAttribute,
  parseOnOffAttribute,
} from "./xmlParser";
import type { XmlElement } from "./xmlParser";

export { formatOoxmlCounter as formatNumber, padDecimal } from "./ooxmlCounterFormatter";

/**
 * Map of rId to numbering definitions
 */
export type NumberingMap = {
  definitions: NumberingDefinitions;
  /** Get level info for a numId and ilvl */
  getLevel: (numId: number, ilvl: number) => ListLevel | null;
  /** Get the abstract numbering ID referenced by a numId */
  getAbstractNumId: (numId: number) => number | null;
  /** Get abstract numbering by ID */
  getAbstract: (abstractNumId: number) => AbstractNumbering | null;
  /** Get the concrete numbering instance for a numId */
  getInstance: (numId: number) => NumberingInstance | null;
  /** Check if numId exists */
  hasNumbering: (numId: number) => boolean;
};

/** `w:abstractNum` reads its id; everything else it carries is the remainder's. */
const MODELLED_ABSTRACT_NUM_ATTRIBUTES: ReadonlySet<string> = new Set(["abstractNumId"]);

/** `w:lvl` reads the level it defines and Word's two template markers. */
const MODELLED_LVL_ATTRIBUTES: ReadonlySet<string> = new Set(["ilvl", "tplc", "tentative"]);

/** `w:num` reads the id paragraphs reference it by. */
const MODELLED_NUM_ATTRIBUTES: ReadonlySet<string> = new Set(["numId"]);

/** `w:lvlOverride` reads the level it overrides. */
const MODELLED_LVL_OVERRIDE_ATTRIBUTES: ReadonlySet<string> = new Set(["ilvl"]);

const attributeValue = (element: XmlElement): string | undefined =>
  getAttribute(element, "w", "val") ?? undefined;

const integerValue = (element: XmlElement): number | undefined => {
  const parsed = parseNumericAttribute(element, "w", "val");
  return parsed === undefined || !Number.isFinite(parsed) ? undefined : parsed;
};

/**
 * A field read off a sequence child, when the child stated one.
 *
 * `CT_AbstractNum` and `CT_Lvl` are sequences and the serializer writes these
 * fields back in that order, so the sink's index is a count of what has been
 * read and a capture lands between the same two neighbours it sat between.
 */
const assignDefined = <Target, Key extends keyof Target>(
  target: Target,
  key: Key,
  carried: Target[Key] | undefined,
): void => {
  if (carried === undefined) {
    return;
  }
  target[key] = carried;
};

const NUMBERING_HANDLERS = {
  abstractNum: (child, definitions) => {
    const parsed = parseAbstractNumbering(child);
    if (parsed) {
      definitions.abstractNums.push(parsed);
    }
  },
  num: (child, definitions) => {
    const parsed = parseNumberingInstance(child);
    if (parsed) {
      definitions.nums.push(parsed);
    }
  },
  // A picture bullet is a whole VML shape or DrawingML drawing that a
  // level names by id, and folio renders the level's `w:lvlText` instead.
  // The bytes go back where they were rather than being rebuilt from a
  // model nothing reads.
  numPicBullet: CAPTURE,
  // Word's high-water mark for the ids it has handed out. Nothing folio
  // does consults it, and minting over it would renumber a document's
  // lists on a machine that never opened it.
  numIdMacAtCleanup: CAPTURE,
} as const satisfies ChildHandlers<"w:numbering", NumberingDefinitions>;

/**
 * Parse numbering.xml into NumberingDefinitions
 *
 * @param numberingXml - Raw XML string from word/numbering.xml (or null if not present)
 * @returns NumberingMap with definitions and helper functions
 */
export function parseNumbering(numberingXml: string | null): NumberingMap {
  const definitions: NumberingDefinitions = {
    abstractNums: [],
    nums: [],
  };

  if (!numberingXml) {
    return createNumberingMap(definitions);
  }

  const root = parseXmlDocument(numberingXml);
  if (!root) {
    return createNumberingMap(definitions);
  }

  const preserved = dispatchChildrenWithContext({
    element: root,
    container: "w:numbering",
    capturePosition: sequencePositions("w:numbering", root),
    handlers: NUMBERING_HANDLERS,
    context: definitions,
  });
  if (preserved !== undefined) {
    definitions.preserved = preserved;
  }
  const preservedAttributes = attributeRemainder({
    element: root,
    modelled: DERIVED_PART_ROOT_ATTRIBUTES,
  });
  if (preservedAttributes !== undefined) {
    definitions.preservedAttributes = preservedAttributes;
  }

  return createNumberingMap(definitions);
}

const ABSTRACT_NUM_HANDLERS = {
  nsid: (child, abstractNum) => assignDefined(abstractNum, "nsid", attributeValue(child)),
  multiLevelType: (child, abstractNum) =>
    assignDefined(abstractNum, "multiLevelType", narrow(attributeValue(child), MULTI_LEVEL_TYPES)),
  tmpl: (child, abstractNum) => assignDefined(abstractNum, "tmpl", attributeValue(child)),
  name: (child, abstractNum) => assignDefined(abstractNum, "name", attributeValue(child)),
  styleLink: (child, abstractNum) => assignDefined(abstractNum, "styleLink", attributeValue(child)),
  numStyleLink: (child, abstractNum) =>
    assignDefined(abstractNum, "numStyleLink", attributeValue(child)),
  lvl: (child, abstractNum) => {
    const level = parseListLevel(child);
    if (level) {
      abstractNum.levels.push(level);
    }
  },
} as const satisfies ChildHandlers<"w:abstractNum", AbstractNumbering>;

/**
 * Parse a single w:abstractNum element
 */
function parseAbstractNumbering(element: XmlElement): AbstractNumbering | null {
  const abstractNumIdStr = getAttribute(element, "w", "abstractNumId");
  if (abstractNumIdStr === null) {
    return null;
  }

  const abstractNumId = Number.parseInt(abstractNumIdStr, 10);
  if (Number.isNaN(abstractNumId)) {
    return null;
  }

  const abstractNum: AbstractNumbering = {
    abstractNumId,
    levels: [],
  };

  const preserved = dispatchChildrenWithContext({
    element,
    container: "w:abstractNum",
    capturePosition: sequencePositions("w:abstractNum", element),
    handlers: ABSTRACT_NUM_HANDLERS,
    context: abstractNum,
  });
  if (preserved !== undefined) {
    abstractNum.preserved = preserved;
  }
  const preservedAttributes = attributeRemainder({
    element,
    modelled: MODELLED_ABSTRACT_NUM_ATTRIBUTES,
  });
  if (preservedAttributes !== undefined) {
    abstractNum.preservedAttributes = preservedAttributes;
  }

  // Sort levels by ilvl
  abstractNum.levels.sort((a, b) => a.ilvl - b.ilvl);

  return abstractNum;
}

/** ECMA-376 §17.9.1: a `w:abstractNum` defines levels 0 through 8 and no more. */
const MAX_LIST_LEVEL = 8;

const MULTI_LEVEL_TYPES = ["hybridMultilevel", "multilevel", "singleLevel"] as const;

const narrow = <Member extends string>(
  carried: string | undefined,
  members: readonly Member[],
): Member | undefined => members.find((member) => member === carried);

/** What a `w:num` walk has read so far. */
type NumberingInstanceWalk = {
  abstractNumId: number | undefined;
  levelOverrides: LevelOverride[];
};

const NUMBERING_INSTANCE_HANDLERS = {
  abstractNumId: (child, walk) => {
    if (walk.abstractNumId !== undefined) {
      return;
    }
    walk.abstractNumId = integerValue(child);
  },
  lvlOverride: (child, { levelOverrides }) => {
    const override = parseLevelOverride(child);
    if (override) {
      levelOverrides.push(override);
    }
  },
} as const satisfies ChildHandlers<"w:num", NumberingInstanceWalk>;

/**
 * Parse a single w:num element (numbering instance)
 */
function parseNumberingInstance(element: XmlElement): NumberingInstance | null {
  const numIdStr = getAttribute(element, "w", "numId");
  if (numIdStr === null) {
    return null;
  }

  const numId = Number.parseInt(numIdStr, 10);
  if (Number.isNaN(numId)) {
    return null;
  }

  const walk: NumberingInstanceWalk = { abstractNumId: undefined, levelOverrides: [] };

  const preserved = dispatchChildrenWithContext({
    element,
    container: "w:num",
    capturePosition: sequencePositions("w:num", element),
    handlers: NUMBERING_INSTANCE_HANDLERS,
    context: walk,
  });
  const { abstractNumId, levelOverrides } = walk;

  // `w:abstractNumId` is the one required child: an instance that names no
  // template numbers nothing, and folio has nowhere to put the rest of it.
  if (abstractNumId === undefined) {
    return null;
  }

  const instance: NumberingInstance = { numId, abstractNumId };
  if (levelOverrides.length > 0) {
    instance.levelOverrides = levelOverrides;
  }
  if (preserved !== undefined) {
    instance.preserved = preserved;
  }
  const preservedAttributes = attributeRemainder({
    element,
    modelled: MODELLED_NUM_ATTRIBUTES,
  });
  if (preservedAttributes !== undefined) {
    instance.preservedAttributes = preservedAttributes;
  }

  return instance;
}

const LEVEL_OVERRIDE_HANDLERS = {
  startOverride: (child, override) => {
    const start = integerValue(child);
    if (start !== undefined && override.startOverride === undefined) {
      override.startOverride = start;
    }
  },
  lvl: (child, override) => {
    const level = parseListLevel(child);
    if (level && override.lvl === undefined) {
      override.lvl = level;
    }
  },
} as const satisfies ChildHandlers<"w:lvlOverride", LevelOverride>;

/**
 * Parse a single w:lvlOverride element (one level of one numbering instance)
 */
function parseLevelOverride(element: XmlElement): LevelOverride | null {
  const ilvl = parseNumericAttribute(element, "w", "ilvl");
  if (ilvl === undefined) {
    return null;
  }

  const override: LevelOverride = { ilvl };

  const preserved = dispatchChildrenWithContext({
    element,
    container: "w:lvlOverride",
    capturePosition: sequencePositions("w:lvlOverride", element),
    handlers: LEVEL_OVERRIDE_HANDLERS,
    context: override,
  });
  if (preserved !== undefined) {
    override.preserved = preserved;
  }
  const preservedAttributes = attributeRemainder({
    element,
    modelled: MODELLED_LVL_OVERRIDE_ATTRIBUTES,
  });
  if (preservedAttributes !== undefined) {
    override.preservedAttributes = preservedAttributes;
  }

  return override;
}

const LIST_LEVEL_HANDLERS = {
  start: (child, level) => assignDefined(level, "start", integerValue(child)),
  numFmt: (child, level) => {
    const resolved = resolveNumFmt(child);
    if (resolved === null) {
      return CAPTURE;
    }
    Object.assign(level, resolved);
    return undefined;
  },
  lvlRestart: (child, level) => assignDefined(level, "lvlRestart", integerValue(child)),
  pStyle: (child, level) => assignDefined(level, "pStyle", attributeValue(child)),
  isLgl: (child, level) => assignDefined(level, "isLgl", parseBooleanElement(child)),
  suff: (child, level) =>
    assignDefined(level, "suffix", narrowEnum(attributeValue(child), LevelSuffixSchema)),
  lvlText: (child, level) => {
    assignDefined(level, "lvlText", attributeValue(child) ?? "");
    const isNull = parseOnOffAttribute(child, "w", "null");
    if (isNull !== undefined) {
      level.lvlTextNull = isNull;
    }
  },
  lvlPicBulletId: (child, level) => assignDefined(level, "lvlPicBulletId", integerValue(child)),
  legacy: (child, level) => assignDefined(level, "legacy", legacyOf(child)),
  lvlJc: (child, level) =>
    assignDefined(level, "lvlJc", narrowEnum(attributeValue(child), ParagraphAlignmentSchema)),
  // An empty `<w:pPr/>` is not an absent one: both of `CT_PPrGeneral`'s
  // children are optional, so the record is what carries the presence.
  pPr: (child, level) => assignDefined(level, "pPr", parseParagraphProperties(child, null) ?? {}),
  rPr: (child, level) =>
    assignDefined(
      level,
      "rPr",
      parseRunProperties(child, null, RUN_PROPERTY_OWNERS.standalone) ?? {},
    ),
} as const satisfies ChildHandlers<"w:lvl", ListLevel>;

/**
 * `mc:AlternateContent` around a custom `w:numFmt` is read for its format by
 * {@link parseListLevel}, and the writer re-emits that format as the level's
 * own `w:numFmt`, so the walk takes it without a capture.
 */
const LIST_LEVEL_UNDECLARED = {
  AlternateContent: () => undefined,
} as const satisfies Record<string, ChildReader<ListLevel>>;

/**
 * Parse a single w:lvl element (list level definition)
 */
function parseListLevel(element: XmlElement): ListLevel | null {
  // A level outside 0..8 names no level a paragraph can reference, but it is
  // still a definition the part carries: the refusal belongs to the lookup in
  // `createNumberingMap`, not to the reader, or the markup is gone.
  const ilvl = parseNumericAttribute(element, "w", "ilvl");
  if (ilvl === undefined) {
    return null;
  }

  const level: ListLevel = {
    ilvl,
    numFmt: "decimal", // Default
    lvlText: "",
  };

  // Word wraps a custom format in mc:AlternateContent — <mc:Choice
  // Requires="w14"> holds <w:numFmt w:val="custom" w:format="..."/> and
  // <mc:Fallback> holds the plain format for pre-w14 readers — so the wrapper
  // is read for its format rather than captured, and the writer re-emits that
  // format as the `w:numFmt` the model holds.
  const alternateEl = findChild(element, "mc", "AlternateContent");

  const preserved = dispatchChildrenWithContext({
    element,
    container: "w:lvl",
    capturePosition: sequencePositions("w:lvl", element),
    handlers: LIST_LEVEL_HANDLERS,
    undeclared: LIST_LEVEL_UNDECLARED,
    context: level,
  });
  if (preserved !== undefined) {
    level.preserved = preserved;
  }
  const preservedAttributes = attributeRemainder({
    element,
    modelled: MODELLED_LVL_ATTRIBUTES,
  });
  if (preservedAttributes !== undefined) {
    level.preservedAttributes = preservedAttributes;
  }

  const tplc = getAttribute(element, "w", "tplc");
  if (tplc !== null) {
    level.tplc = tplc;
  }
  const tentative = parseOnOffAttribute(element, "w", "tentative");
  if (tentative !== undefined) {
    level.tentative = tentative;
  }

  // ECMA-376 Part 3 §10.2.1: a consumer that does not understand a Choice
  // takes the Fallback. Prefer the Choice when its format resolves to
  // something folio implements; otherwise the Fallback is the closer
  // rendering.
  if (alternateEl && findChild(element, "w", "numFmt") === null) {
    const choiceFmt = resolveNumFmt(
      findChild(findChild(alternateEl, "mc", "Choice"), "w", "numFmt"),
    );
    const fallbackFmt = resolveNumFmt(
      findChild(findChild(alternateEl, "mc", "Fallback"), "w", "numFmt"),
    );
    // ECMA-376 Part 3 §10.2.1: a consumer that does not understand the Choice
    // takes the Fallback. folio understands `custom` well enough to save it,
    // but not to count in it unless its format is a zero-padded decimal, so an
    // unrenderable Choice defers to the Fallback. The Choice is not lost by
    // that: the whole `mc:AlternateContent` replays verbatim on save.
    const resolved =
      choiceFmt && counterFormatOf(choiceFmt) !== "custom" ? choiceFmt : (fallbackFmt ?? choiceFmt);
    if (resolved) {
      Object.assign(level, resolved);
    }
  }

  return level;
}

/** A `w:numFmt` as the model holds it: its token and authored format metadata. */
type ResolvedNumFmt = Pick<ListLevel, "numFmt" | "numFmtFormat">;

/**
 * Resolve a `<w:numFmt>` element, or null when it is absent or carries a token
 * outside `ST_NumberFormat` (an mc:Fallback can then supply the rendering).
 *
 * `@w:format` stays beside every admitted token exactly as authored. `custom`
 * used to be decoded here into a synthetic `decimalZero{3,4,5}` the
 * enumeration does not declare, which the serializer wrote back as a `w:val`
 * no consumer can read. The pad-width interpretation belongs to the renderer,
 * not to the model; preservation belongs to every schema-valid `CT_NumFmt`.
 */
function resolveNumFmt(numFmtEl: XmlElement | null): ResolvedNumFmt | null {
  if (!numFmtEl) {
    return null;
  }
  const numFmt = narrowEnum(getAttribute(numFmtEl, "w", "val"), NumberFormatSchema);
  if (numFmt === undefined) {
    return null;
  }
  const format = getAttribute(numFmtEl, "w", "format");
  return format === null ? { numFmt } : { numFmt, numFmtFormat: format };
}

/**
 * `w:legacy`: the flag is the element's own `w:legacy` attribute.
 *
 * `CT_LvlLegacy` has no `w:val`, so reading one turned every explicit
 * `w:legacy="0"` into the `w:legacy="1"` a rebuild wrote back.
 */
function legacyOf(element: XmlElement): LevelLegacy {
  const legacy = parseOnOffAttribute(element, "w", "legacy");
  const legacySpace = parseNumericAttribute(element, "w", "legacySpace");
  const legacyIndent = parseNumericAttribute(element, "w", "legacyIndent");
  return {
    ...(legacy === undefined ? {} : { legacy }),
    ...(legacySpace === undefined ? {} : { legacySpace }),
    ...(legacyIndent === undefined ? {} : { legacyIndent }),
  };
}

/**
 * The width a `custom` format's zero-padded first token counts to, or
 * `undefined` for a pattern folio does not render.
 *
 * ECMA-376 §17.9.17: `@w:format` is an XSLT token list like
 * "0001, 0002, 0003, ...". Word only emits zero-padded decimal customs this
 * way; the width is the digit count of the first token, clamped to five.
 */
export const customNumberFormatPadWidth = (format: string | undefined): number | undefined => {
  const firstToken = format?.split(",")[0]?.trim() ?? "";
  if (!/^0+1$/u.test(firstToken)) {
    return undefined;
  }
  return Math.min(firstToken.length, 5);
};

/** What the marker renderer counts a level in; `custom` decides by pad width. */
export const counterFormatOf = (level: ResolvedNumFmt): CounterFormat => {
  if (level.numFmt !== "custom") {
    return level.numFmt;
  }
  switch (customNumberFormatPadWidth(level.numFmtFormat)) {
    case 2:
      return "decimalZero";
    case 3:
      return "decimalZero3";
    case 4:
      return "decimalZero4";
    case 5:
      return "decimalZero5";
    default:
      return "custom";
  }
};

/**
 * How a `w:lvlJc` lays a marker out, for the three alignments layout has.
 *
 * `start` and `end` are the Strict logical-direction spellings folio
 * canonicalises to `left` and `right` everywhere it rebuilds a part. The
 * justification members describe how a *line* of text is spread and say nothing
 * about a marker, so they resolve to no alignment rather than to a guess; the
 * value itself is still carried on the level and written back.
 */
const MARKER_ALIGNMENT = {
  start: "left",
  left: "left",
  center: "center",
  end: "right",
  right: "right",
  both: undefined,
  distribute: undefined,
  mediumKashida: undefined,
  highKashida: undefined,
  lowKashida: undefined,
  numTab: undefined,
  thaiDistribute: undefined,
} as const satisfies Record<ParagraphAlignment, ListRendering["markerAlignment"]>;

/** {@link MARKER_ALIGNMENT} for a level that may not state a justification. */
export const markerAlignmentForLevel = (
  lvlJc: ParagraphAlignment | undefined,
): ListRendering["markerAlignment"] => (lvlJc === undefined ? undefined : MARKER_ALIGNMENT[lvlJc]);

export const markerFormattingFromLevel = (
  formatting: TextFormatting | undefined,
): ListMarkerFormatting | undefined => {
  if (!formatting) {
    return undefined;
  }
  const markerFormatting: ListMarkerFormatting = {
    ...(formatting.fontFamily !== undefined ? { fontFamily: formatting.fontFamily } : {}),
    ...(formatting.fontSize !== undefined ? { fontSize: formatting.fontSize } : {}),
    ...(formatting.fontSizeCs !== undefined ? { fontSizeCs: formatting.fontSizeCs } : {}),
    ...(formatting.bold !== undefined ? { bold: formatting.bold } : {}),
    ...(formatting.boldCs !== undefined ? { boldCs: formatting.boldCs } : {}),
    ...(formatting.italic !== undefined ? { italic: formatting.italic } : {}),
    ...(formatting.italicCs !== undefined ? { italicCs: formatting.italicCs } : {}),
    ...(formatting.rtl !== undefined ? { rtl: formatting.rtl } : {}),
    ...(formatting.cs !== undefined ? { cs: formatting.cs } : {}),
  };
  return Object.keys(markerFormatting).length > 0 ? markerFormatting : undefined;
};

/**
 * Per-definitions cache for `createNumberingMap`. Style application rebuilds
 * the lookup map on every picker click otherwise; the definitions object is
 * stable for a document's lifetime, so cache by identity.
 */
const numberingMapCache = new WeakMap<NumberingDefinitions, NumberingMap>();

export function getCachedNumberingMap(definitions: NumberingDefinitions): NumberingMap {
  let map = numberingMapCache.get(definitions);
  if (!map) {
    map = createNumberingMap(definitions);
    numberingMapCache.set(definitions, map);
  }
  return map;
}

/**
 * Create a NumberingMap with helper functions
 */
export function createNumberingMap(definitions: NumberingDefinitions): NumberingMap {
  // Build lookup maps for efficient access
  const abstractMap = new Map<number, AbstractNumbering>();
  for (const abs of definitions.abstractNums) {
    abstractMap.set(abs.abstractNumId, abs);
  }

  const numMap = new Map<number, NumberingInstance>();
  for (const num of definitions.nums) {
    numMap.set(num.numId, num);
  }

  return {
    definitions,

    getLevel(numId: number, ilvl: number): ListLevel | null {
      // A `w:abstractNum` defines at most nine levels, so an `ilvl` outside 0
      // through 8 names none however the part spells it. The refusal lives
      // here rather than in the reader so a definition folio cannot resolve is
      // still a definition it carries.
      if (ilvl < 0 || ilvl > MAX_LIST_LEVEL) {
        return null;
      }
      const num = numMap.get(numId);
      if (!num) {
        return null;
      }

      // Check for level override first
      if (num.levelOverrides) {
        const override = num.levelOverrides.find((o) => o.ilvl === ilvl);
        if (override) {
          if (override.lvl) {
            // Full level redefinition
            return override.lvl;
          }
          // Start override - need to get base level and modify
          const abstractNum = abstractMap.get(num.abstractNumId);
          if (abstractNum) {
            const baseLevel = abstractNum.levels.find((l) => l.ilvl === ilvl);
            if (baseLevel && override.startOverride !== undefined) {
              return {
                ...baseLevel,
                start: override.startOverride,
              };
            }
          }
        }
      }

      // Get from abstract numbering
      let abstractNum = abstractMap.get(num.abstractNumId);
      if (!abstractNum) {
        return null;
      }

      // Follow numStyleLink: when an abstractNum has numStyleLink instead of
      // defining levels directly, find the abstractNum that owns that style
      // (has matching styleLink) and use its levels. Per ECMA-376 §17.9.21/22.
      if (abstractNum.numStyleLink && abstractNum.levels.length === 0) {
        for (const candidate of abstractMap.values()) {
          if (candidate.styleLink === abstractNum.numStyleLink && candidate.levels.length > 0) {
            abstractNum = candidate;
            break;
          }
        }
      }

      return abstractNum.levels.find((l) => l.ilvl === ilvl) ?? null;
    },

    getAbstractNumId(numId: number): number | null {
      return numMap.get(numId)?.abstractNumId ?? null;
    },

    getAbstract(abstractNumId: number): AbstractNumbering | null {
      return abstractMap.get(abstractNumId) ?? null;
    },

    getInstance(numId: number): NumberingInstance | null {
      return numMap.get(numId) ?? null;
    },

    hasNumbering(numId: number): boolean {
      return numMap.has(numId);
    },
  };
}

/**
 * Resolve a paragraph's `numPr` against the numbering definitions into the
 * base `ListRendering` the layout pipeline needs (marker template, per-level
 * numFmts, counter key, start override). Returns null when the numPr doesn't
 * name a real level — including `numId === 0`, "no numbering" per ECMA-376.
 *
 * Shared by the style picker (`listAttrsFromResolvedStyle`) so a style-attached
 * list renders the same marker the document loader produces. The loader's
 * `parseParagraph` block additionally folds inline LISTNUM markers and computes
 * the second-slot offset from the paragraph's content; those fields are absent
 * here because the picker has no paragraph content to inspect.
 */
export function computeListRendering(
  numPr: { numId?: number; ilvl?: number },
  numbering: NumberingMap,
): ListRendering | null {
  const { numId, ilvl = 0 } = numPr;
  if (!isNumberingReference(numId)) {
    return null;
  }

  const level = numbering.getLevel(numId, ilvl);
  if (!level) {
    return null;
  }

  // Collect numFmts for levels 0..ilvl so multi-level templates like "%1.%2."
  // can resolve each %N with its own format (legal numbering forces decimal).
  const levelNumFmts: CounterFormat[] = [];
  const levelStarts: number[] = [];
  for (let i = 0; i <= ilvl; i += 1) {
    const listLevel = numbering.getLevel(numId, i);
    levelNumFmts.push(level.isLgl || !listLevel ? "decimal" : counterFormatOf(listLevel));
    levelStarts.push(listLevel?.start ?? 1);
  }

  const instance = numbering.getInstance(numId);
  const overrideForLevel = instance?.levelOverrides?.find((override) => override.ilvl === ilvl);

  const rendering: ListRendering = {
    level: ilvl,
    numId,
    marker: level.lvlText,
    markerTemplate: level.lvlText,
    isBullet: level.numFmt === "bullet",
    numFmt: level.isLgl ? "decimal" : counterFormatOf(level),
    levelNumFmts,
    levelStarts,
  };
  if (level.isLgl) {
    rendering.isLegal = true;
  }
  if (level.rPr?.hidden) {
    rendering.markerHidden = true;
  }
  const markerFormatting = markerFormattingFromLevel(level.rPr);
  if (markerFormatting) {
    rendering.markerFormatting = markerFormatting;
  }
  if (level.rPr?.allCaps) {
    rendering.markerAllCaps = true;
  }
  const markerAlignment = markerAlignmentForLevel(level.lvlJc);
  if (markerAlignment !== undefined) {
    rendering.markerAlignment = markerAlignment;
  }
  if (level.suffix) {
    rendering.markerSuffix = level.suffix;
  }
  if (instance?.abstractNumId !== undefined) {
    rendering.abstractNumId = instance.abstractNumId;
  }
  if (overrideForLevel?.startOverride !== undefined) {
    rendering.startOverride = overrideForLevel.startOverride;
  }
  return rendering;
}

/**
 * Render list marker text by replacing placeholders with formatted numbers
 *
 * @param lvlText - The level text pattern (e.g., "%1.", "%1.%2")
 * @param counters - Array of counter values for each level (index 0 = level 0, etc.)
 * @param formats - Array of number formats for each level
 * @returns Rendered marker text
 */
export function renderListMarker(
  lvlText: string,
  counters: number[],
  formats: CounterFormat[],
): string {
  let result = lvlText;

  // Replace %1 through %9 with formatted counter values
  for (let i = 1; i <= 9; i++) {
    const placeholder = `%${i}`;
    if (result.includes(placeholder)) {
      const counterIndex = i - 1;
      const counter = counters[counterIndex] ?? 1;
      const format = formats[counterIndex] ?? "decimal";
      const formatted = formatOoxmlCounter(counter, format);
      result = result.replaceAll(placeholder, formatted);
    }
  }

  return result;
}

/**
 * Get the bullet character for a bullet list level
 *
 * @param level - The list level definition
 * @returns The bullet character to display
 */
export function getBulletCharacter(level: ListLevel): string {
  // If lvlText is set and not empty, use it
  if (level.lvlText) {
    return level.lvlText;
  }

  // Check font for common bullet font mappings
  const fontFamily = level.rPr?.fontFamily?.ascii || level.rPr?.fontFamily?.hAnsi;

  if (fontFamily) {
    const fontLower = fontFamily.toLowerCase();

    // Symbol font common bullets
    if (fontLower === "symbol") {
      return "•"; // Standard bullet
    }

    // Wingdings common bullets
    if (fontLower.includes("wingding")) {
      return "❑"; // Square bullet
    }
  }

  // Default bullet
  return "•";
}

/**
 * Check if a list level is a bullet (not numbered)
 */
export function isBulletLevel(level: ListLevel): boolean {
  return level.numFmt === "bullet" || level.numFmt === "none";
}

/** Whether a numbering level reserves horizontal space for a visible marker. */
export const numberingLevelHasMarkerSlot = (level: Pick<ListLevel, "numFmt">): boolean =>
  level.numFmt !== "none";
