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
  NumberingInstance,
  ListLevel,
  ListRendering,
  CounterFormat,
  NumberFormat,
  ParagraphFormatting,
  TextFormatting,
  ListMarkerFormatting,
} from "../types/document";
import { isNumberingReference } from "./numberingReference";
import { formatOoxmlCounter } from "./ooxmlCounterFormatter";
import {
  LevelSuffixSchema,
  narrowEnum,
  NumberFormatSchema,
  TabLeaderSchema,
  TabStopAlignmentSchema,
} from "./parserEnums";
import { parseRunProperties } from "./runParser";
import {
  parseXmlDocument,
  findChild,
  findChildren,
  getAttribute,
  getLocalName,
  getNamespaceUri,
  parseBooleanElement,
  parseNumericAttribute,
  WORDPROCESSINGML_NAMESPACE_URIS,
} from "./xmlParser";
import type { XmlElement } from "./xmlParser";
import { numericAttributeAnySpelling } from "./strictNames";

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

const wordprocessingLocalName = (element: XmlElement): string | null =>
  WORDPROCESSINGML_NAMESPACE_URIS.has(getNamespaceUri(element) ?? "")
    ? getLocalName(element.name)
    : null;

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

  // Parse abstract numbering definitions
  const abstractNumElements = findChildren(root, "w", "abstractNum");
  for (const abstractNum of abstractNumElements) {
    const parsed = parseAbstractNumbering(abstractNum);
    if (parsed) {
      definitions.abstractNums.push(parsed);
    }
  }

  // Parse numbering instances
  const numElements = findChildren(root, "w", "num");
  for (const num of numElements) {
    const parsed = parseNumberingInstance(num);
    if (parsed) {
      definitions.nums.push(parsed);
    }
  }

  return createNumberingMap(definitions);
}

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

  let multiLevelTypeEl: XmlElement | undefined;
  let nameEl: XmlElement | undefined;
  let numStyleLinkEl: XmlElement | undefined;
  let styleLinkEl: XmlElement | undefined;
  const levelElements: XmlElement[] = [];

  for (const child of element.elements ?? []) {
    if (child.type !== "element") {
      continue;
    }

    switch (wordprocessingLocalName(child)) {
      case "multiLevelType":
        multiLevelTypeEl ??= child;
        break;
      case "name":
        nameEl ??= child;
        break;
      case "numStyleLink":
        numStyleLinkEl ??= child;
        break;
      case "styleLink":
        styleLinkEl ??= child;
        break;
      case "lvl":
        levelElements.push(child);
        break;
      default:
        break;
    }
  }

  if (multiLevelTypeEl) {
    const mlType = getAttribute(multiLevelTypeEl, "w", "val");
    if (mlType === "hybridMultilevel" || mlType === "multilevel" || mlType === "singleLevel") {
      abstractNum.multiLevelType = mlType;
    }
  }

  // Parse name
  if (nameEl) {
    const nameVal = getAttribute(nameEl, "w", "val");
    if (nameVal != null) {
      abstractNum.name = nameVal;
    }
  }

  // Parse style links
  if (numStyleLinkEl) {
    const numStyleLinkVal = getAttribute(numStyleLinkEl, "w", "val");
    if (numStyleLinkVal != null) {
      abstractNum.numStyleLink = numStyleLinkVal;
    }
  }

  if (styleLinkEl) {
    const styleLinkVal = getAttribute(styleLinkEl, "w", "val");
    if (styleLinkVal != null) {
      abstractNum.styleLink = styleLinkVal;
    }
  }

  // Parse levels (w:lvl)
  for (const lvlEl of levelElements) {
    const level = parseListLevel(lvlEl);
    if (level) {
      abstractNum.levels.push(level);
    }
  }

  // Sort levels by ilvl
  abstractNum.levels.sort((a, b) => a.ilvl - b.ilvl);

  return abstractNum;
}

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

  let abstractNumIdEl: XmlElement | undefined;
  const overrideElements: XmlElement[] = [];
  for (const child of element.elements ?? []) {
    if (child.type !== "element") {
      continue;
    }

    if (wordprocessingLocalName(child) === "abstractNumId") {
      abstractNumIdEl ??= child;
      continue;
    }

    if (wordprocessingLocalName(child) === "lvlOverride") {
      overrideElements.push(child);
    }
  }

  if (!abstractNumIdEl) {
    return null;
  }

  const abstractNumIdStr = getAttribute(abstractNumIdEl, "w", "val");
  if (abstractNumIdStr === null) {
    return null;
  }

  const abstractNumId = Number.parseInt(abstractNumIdStr, 10);
  if (Number.isNaN(abstractNumId)) {
    return null;
  }

  const instance: NumberingInstance = {
    numId,
    abstractNumId,
  };

  // Parse level overrides (w:lvlOverride)
  if (overrideElements.length > 0) {
    instance.levelOverrides = [];

    for (const overrideEl of overrideElements) {
      const ilvlStr = getAttribute(overrideEl, "w", "ilvl");
      if (ilvlStr === null) {
        continue;
      }

      const ilvl = Number.parseInt(ilvlStr, 10);
      if (Number.isNaN(ilvl)) {
        continue;
      }

      const override: {
        ilvl: number;
        startOverride?: number;
        lvl?: ListLevel;
      } = { ilvl };

      let startOverrideEl: XmlElement | undefined;
      let lvlEl: XmlElement | undefined;
      for (const child of overrideEl.elements ?? []) {
        if (child.type !== "element") {
          continue;
        }

        if (wordprocessingLocalName(child) === "startOverride") {
          startOverrideEl ??= child;
          continue;
        }

        if (wordprocessingLocalName(child) === "lvl") {
          lvlEl ??= child;
        }
      }

      // Check for start override
      if (startOverrideEl) {
        const startVal = getAttribute(startOverrideEl, "w", "val");
        if (startVal !== null) {
          const startNum = Number.parseInt(startVal, 10);
          if (!Number.isNaN(startNum)) {
            override.startOverride = startNum;
          }
        }
      }

      // Check for full level redefinition
      if (lvlEl) {
        const parsedLvl = parseListLevel(lvlEl);
        if (parsedLvl != null) {
          override.lvl = parsedLvl;
        }
      }

      instance.levelOverrides.push(override);
    }
  }

  return instance;
}

/**
 * Parse a single w:lvl element (list level definition)
 */
function parseListLevel(element: XmlElement): ListLevel | null {
  const ilvlStr = getAttribute(element, "w", "ilvl");
  if (ilvlStr === null) {
    return null;
  }

  const ilvl = Number.parseInt(ilvlStr, 10);
  if (Number.isNaN(ilvl) || ilvl < 0 || ilvl > 8) {
    return null;
  }

  const level: ListLevel = {
    ilvl,
    numFmt: "decimal", // Default
    lvlText: "",
  };

  let startEl: XmlElement | undefined;
  let numFmtEl: XmlElement | undefined;
  let alternateEl: XmlElement | undefined;
  let lvlTextEl: XmlElement | undefined;
  let lvlJcEl: XmlElement | undefined;
  let suffEl: XmlElement | undefined;
  let isLglEl: XmlElement | undefined;
  let lvlRestartEl: XmlElement | undefined;
  let legacyEl: XmlElement | undefined;
  let pPrEl: XmlElement | undefined;
  let rPrEl: XmlElement | undefined;

  for (const child of element.elements ?? []) {
    if (child.type !== "element") {
      continue;
    }

    const localName = wordprocessingLocalName(child);
    if (localName === null) {
      if (getLocalName(child.name) === "AlternateContent") {
        alternateEl ??= child;
      }
      continue;
    }
    switch (localName) {
      case "start":
        startEl ??= child;
        break;
      case "numFmt":
        numFmtEl ??= child;
        break;
      case "lvlText":
        lvlTextEl ??= child;
        break;
      case "lvlJc":
        lvlJcEl ??= child;
        break;
      case "suff":
        suffEl ??= child;
        break;
      case "isLgl":
        isLglEl ??= child;
        break;
      case "lvlRestart":
        lvlRestartEl ??= child;
        break;
      case "legacy":
        legacyEl ??= child;
        break;
      case "pPr":
        pPrEl ??= child;
        break;
      case "rPr":
        rPrEl ??= child;
        break;
      default:
        break;
    }
  }

  // Parse start value
  if (startEl) {
    const startVal = getAttribute(startEl, "w", "val");
    if (startVal !== null) {
      const startNum = Number.parseInt(startVal, 10);
      if (!Number.isNaN(startNum)) {
        level.start = startNum;
      }
    }
  }

  // Parse number format. Word wraps custom formats in mc:AlternateContent —
  // <mc:Choice Requires="w14"> holds <w:numFmt w:val="custom" w:format="..."/>
  // and <mc:Fallback> holds the plain format for pre-w14 readers. Prefer the
  // Choice when its format resolves to something we implement; otherwise the
  // Fallback is the closer rendering (ECMA-376 Part 3 §10.2.1 — a consumer
  // that doesn't understand a Choice must take the Fallback).
  if (numFmtEl) {
    Object.assign(level, resolveNumFmt(numFmtEl) ?? { numFmt: "decimal" });
  } else if (alternateEl) {
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

  // Parse level text (the pattern like "%1." or "•")
  if (lvlTextEl) {
    level.lvlText = getAttribute(lvlTextEl, "w", "val") ?? "";
  }

  // Parse justification
  if (lvlJcEl) {
    const jcVal = getAttribute(lvlJcEl, "w", "val");
    if (jcVal === "left" || jcVal === "center" || jcVal === "right") {
      level.lvlJc = jcVal;
    }
  }

  // Parse suffix
  if (suffEl) {
    const suffix = narrowEnum(getAttribute(suffEl, "w", "val"), LevelSuffixSchema);
    if (suffix) {
      level.suffix = suffix;
    }
  }

  // Parse isLgl (legal numbering)
  if (isLglEl) {
    level.isLgl = parseBooleanElement(isLglEl);
  }

  // Parse lvlRestart (restart numbering from a higher level)
  if (lvlRestartEl) {
    const restartVal = getAttribute(lvlRestartEl, "w", "val");
    if (restartVal !== null) {
      const restartNum = Number.parseInt(restartVal, 10);
      if (!Number.isNaN(restartNum)) {
        level.lvlRestart = restartNum;
      }
    }
  }

  // Parse legacy settings
  if (legacyEl) {
    const legacySpace = parseNumericAttribute(legacyEl, "w", "legacySpace");
    const legacyIndent = parseNumericAttribute(legacyEl, "w", "legacyIndent");
    level.legacy = {
      legacy: parseBooleanElement(legacyEl),
      ...(legacySpace !== undefined ? { legacySpace } : {}),
      ...(legacyIndent !== undefined ? { legacyIndent } : {}),
    };
  }

  // Parse paragraph properties (w:pPr)
  if (pPrEl) {
    level.pPr = parseLevelParagraphProps(pPrEl);
  }

  // Parse run properties (w:rPr)
  if (rPrEl) {
    const runProperties = parseRunProperties(rPrEl, null);
    if (runProperties) {
      level.rPr = runProperties;
    }
  }

  return level;
}

/** A `w:numFmt` as the model holds it: the token, plus what `custom` defers to. */
type ResolvedNumFmt = Pick<ListLevel, "numFmt" | "numFmtFormat">;

/**
 * Resolve a `<w:numFmt>` element, or null when it is absent or carries a token
 * outside `ST_NumberFormat` (an mc:Fallback can then supply the rendering).
 *
 * `custom` stays `custom`, with its `@w:format` beside it. It used to be
 * decoded here into a synthetic `decimalZero{3,4,5}` the enumeration does not
 * declare, which the serializer wrote back as a `w:val` no consumer can read.
 * The pad width belongs to the renderer, not to the model.
 */
function resolveNumFmt(numFmtEl: XmlElement | null): ResolvedNumFmt | null {
  if (!numFmtEl) {
    return null;
  }
  const numFmt = narrowEnum(getAttribute(numFmtEl, "w", "val"), NumberFormatSchema);
  if (numFmt === undefined) {
    return null;
  }
  if (numFmt !== "custom") {
    return { numFmt };
  }
  const format = getAttribute(numFmtEl, "w", "format");
  return format === null ? { numFmt } : { numFmt, numFmtFormat: format };
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

/**
 * The `@w:format` a level's `w:numFmt` writes, which only `custom` carries.
 *
 * `custom` is the reserved member that defers to another attribute, so the
 * comparison lives here, in the module the reserved-value registry names as
 * its owner, rather than in the serializer that needs the answer.
 */
export const customNumberFormatOf = (level: ResolvedNumFmt): string | undefined =>
  level.numFmt === "custom" ? level.numFmtFormat : undefined;

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
 * Parse paragraph properties for a list level (subset of full pPr)
 * Main concern: indentation and tabs
 */
function parseLevelParagraphProps(pPr: XmlElement): ParagraphFormatting {
  const formatting: ParagraphFormatting = {};

  let indEl: XmlElement | undefined;
  let tabsEl: XmlElement | undefined;
  for (const child of pPr.elements ?? []) {
    if (child.type !== "element") {
      continue;
    }

    if (wordprocessingLocalName(child) === "ind") {
      indEl ??= child;
      continue;
    }

    if (wordprocessingLocalName(child) === "tabs") {
      tabsEl ??= child;
    }
  }

  // Parse indentation (w:ind)
  if (indEl) {
    const left = numericAttributeAnySpelling(indEl, "CT_Ind @left");
    const right = numericAttributeAnySpelling(indEl, "CT_Ind @right");
    const firstLine = parseNumericAttribute(indEl, "w", "firstLine");
    const hanging = parseNumericAttribute(indEl, "w", "hanging");

    if (left !== undefined) {
      formatting.indentLeft = left;
    }
    if (right !== undefined) {
      formatting.indentRight = right;
    }

    if (hanging !== undefined) {
      formatting.indentFirstLine = -hanging;
      formatting.hangingIndent = true;
    } else if (firstLine !== undefined) {
      formatting.indentFirstLine = firstLine;
    }
  }

  // Parse tabs (w:tabs)
  if (tabsEl) {
    formatting.tabs = [];
    const tabElements = findChildren(tabsEl, "w", "tab");
    for (const tabEl of tabElements) {
      const pos = parseNumericAttribute(tabEl, "w", "pos");
      const val = getAttribute(tabEl, "w", "val");
      const leader = getAttribute(tabEl, "w", "leader");

      const alignment = narrowEnum(val, TabStopAlignmentSchema);
      if (pos !== undefined && alignment) {
        const parsedLeader = narrowEnum(leader, TabLeaderSchema);
        formatting.tabs.push({
          position: pos,
          alignment,
          ...(parsedLeader !== undefined ? { leader: parsedLeader } : {}),
        });
      }
    }
  }

  return formatting;
}

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
  if (level.lvlJc) {
    rendering.markerAlignment = level.lvlJc;
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
  formats: NumberFormat[],
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
