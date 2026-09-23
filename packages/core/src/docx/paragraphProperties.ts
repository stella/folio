/**
 * The one reader for a paragraph property set (`w:pPr`).
 *
 * Four owners produce the element — a paragraph, the `CT_PPrGeneral` a style
 * carries, a numbering level's, and the `CT_PPrBase` snapshot inside
 * `w:pPrChange` — and the writer they share is
 * `serializeParagraphPropertySet`. The reader lives here rather than beside
 * one of them because `paragraphParser` and `numberingParser` already depend
 * on each other: a reader inside either is one the other cannot import, and a
 * numbering level kept a third private copy — indent and tabs, no sink — for
 * exactly that reason. Nothing in this module imports either of them, so both
 * import it.
 */

import type { ParagraphFormatting, TabStop, Theme } from "../types/document";
import { outlineLevelFromStatedValue } from "@stll/docx-core/model";
import { readAttributeBag } from "./attributeRemainder";
import { parseBorderSpec } from "./borderParser";
import {
  CAPTURE,
  type ChildHandlers,
  type ChildReader,
  dispatchChildrenWithContext,
  ownedElsewhere,
  sequencePositions,
} from "./containerChildren";
import { readParagraphNumbering } from "./numberingReference";
import {
  FrameWrapSchema,
  FrameXAlignSchema,
  FrameYAlignSchema,
  LineSpacingRuleSchema,
  narrowEnum,
  ParagraphAlignmentSchema,
  TabLeaderSchema,
  TabStopAlignmentSchema,
} from "./parserEnums";
import {
  FRAME_ATTRIBUTES,
  INDENTATION_ATTRIBUTES,
  SPACING_ATTRIBUTES,
  TAB_STOP_ATTRIBUTES,
} from "./propertyElementAttributes";
import { parseRunProperties, RUN_PROPERTY_OWNERS } from "./runParser";
import { parseShading } from "./shadingParser";
import { numericAttributeAnySpelling } from "./strictNames";
import { captureVerbatimXml } from "./verbatimCapture";
import {
  findChild,
  findChildren,
  getAttribute,
  parseBooleanElement,
  parseNumericAttribute,
  parseOnOffAttribute,
  parseOnOffChild,
} from "./xmlParser";
import type { XmlElement } from "./xmlParser";

/**
 * Parse tab stops (w:tabs)
 */
function parseTabStops(tabs: XmlElement | null): TabStop[] | undefined {
  if (!tabs) {
    return undefined;
  }

  const tabElements = findChildren(tabs, "w", "tab");
  if (tabElements.length === 0) {
    return undefined;
  }

  const result: TabStop[] = [];

  for (const tab of tabElements) {
    const pos = parseNumericAttribute(tab, "w", "pos");
    const alignment = narrowEnum(getAttribute(tab, "w", "val"), TabStopAlignmentSchema);

    if (pos !== undefined && alignment) {
      const tabStop: TabStop = {
        position: pos,
        alignment,
      };

      const leader = narrowEnum(getAttribute(tab, "w", "leader"), TabLeaderSchema);
      if (leader && leader !== "none") {
        tabStop.leader = leader;
      }

      const preservedAttributes = readAttributeBag(tab, TAB_STOP_ATTRIBUTES);
      if (preservedAttributes) {
        tabStop.preservedAttributes = preservedAttributes;
      }

      result.push(tabStop);
    }
  }

  return result.length > 0 ? result : undefined;
}

/**
 * Parse frame properties (w:framePr)
 */
function parseFrameProperties(
  framePr: XmlElement | null,
): ParagraphFormatting["frame"] | undefined {
  if (!framePr) {
    return undefined;
  }

  const frame: ParagraphFormatting["frame"] = {};

  const dropCap = getAttribute(framePr, "w", "dropCap");
  if (dropCap === "none" || dropCap === "drop" || dropCap === "margin") {
    frame.dropCap = dropCap;
  }

  const lines = parseNumericAttribute(framePr, "w", "lines");
  if (lines !== undefined) {
    frame.lines = lines;
  }

  const w = parseNumericAttribute(framePr, "w", "w");
  if (w !== undefined) {
    frame.width = w;
  }

  const h = parseNumericAttribute(framePr, "w", "h");
  if (h !== undefined) {
    frame.height = h;
  }

  const hSpace = parseNumericAttribute(framePr, "w", "hSpace");
  if (hSpace !== undefined) {
    frame.hSpace = hSpace;
  }

  const vSpace = parseNumericAttribute(framePr, "w", "vSpace");
  if (vSpace !== undefined) {
    frame.vSpace = vSpace;
  }

  const hAnchor = getAttribute(framePr, "w", "hAnchor");
  if (hAnchor === "text" || hAnchor === "margin" || hAnchor === "page") {
    frame.hAnchor = hAnchor;
  }

  const vAnchor = getAttribute(framePr, "w", "vAnchor");
  if (vAnchor === "text" || vAnchor === "margin" || vAnchor === "page") {
    frame.vAnchor = vAnchor;
  }

  const x = parseNumericAttribute(framePr, "w", "x");
  if (x !== undefined) {
    frame.x = x;
  }

  const y = parseNumericAttribute(framePr, "w", "y");
  if (y !== undefined) {
    frame.y = y;
  }

  const xAlign = narrowEnum(getAttribute(framePr, "w", "xAlign"), FrameXAlignSchema);
  if (xAlign) {
    frame.xAlign = xAlign;
  }

  const yAlign = narrowEnum(getAttribute(framePr, "w", "yAlign"), FrameYAlignSchema);
  if (yAlign) {
    frame.yAlign = yAlign;
  }

  const wrap = narrowEnum(getAttribute(framePr, "w", "wrap"), FrameWrapSchema);
  if (wrap) {
    frame.wrap = wrap;
  }

  // Only when the element is modelled at all: a `w:framePr` folio takes
  // nothing from is handed back to the dispatcher and kept whole, and a
  // remainder as well would write the same attributes twice.
  if (Object.keys(frame).length === 0) {
    return undefined;
  }
  const preservedAttributes = readAttributeBag(framePr, FRAME_ATTRIBUTES);
  if (preservedAttributes) {
    frame.preservedAttributes = preservedAttributes;
  }

  return frame;
}

/** A `w:pPr` child folio models as a tri-state `CT_OnOff` toggle. */
type ParagraphToggleField = {
  [Field in keyof ParagraphFormatting]-?: NonNullable<ParagraphFormatting[Field]> extends boolean
    ? Field
    : never;
}[keyof ParagraphFormatting];

/** What a `w:pPr` handler reads into, and the properties already taken. */
type ParagraphPropertyReadContext = {
  formatting: ParagraphFormatting;
  taken: Set<string>;
};

type ParagraphPropertyReader = ChildReader<ParagraphPropertyReadContext>;

/**
 * The content model declares each property once. A source that states one
 * twice keeps the first as the model's value and the repeat as bytes at the
 * same schema ordinal, rather than letting the second silently win or fall off
 * the end of the walk.
 */
const readOnce =
  (
    name: string,
    read: (child: XmlElement, formatting: ParagraphFormatting) => boolean,
  ): ParagraphPropertyReader =>
  (child, { formatting, taken }) => {
    if (taken.has(name) || !read(child, formatting)) {
      return CAPTURE;
    }
    taken.add(name);
    return undefined;
  };

const readToggle = (name: string, field: ParagraphToggleField): ParagraphPropertyReader =>
  readOnce(name, (child, formatting) => {
    formatting[field] = parseBooleanElement(child);
    return true;
  });

/**
 * The `w:pPr` decision map. Built once: every reader takes the record it fills
 * from the walk's context rather than closing over it.
 */
const PARAGRAPH_PROPERTY_HANDLERS = {
  pStyle: readOnce("pStyle", (child, formatting) => {
    const val = getAttribute(child, "w", "val");
    if (!val) {
      return false;
    }
    formatting.styleId = val;
    return true;
  }),
  keepNext: readToggle("keepNext", "keepNext"),
  keepLines: readToggle("keepLines", "keepLines"),
  pageBreakBefore: readToggle("pageBreakBefore", "pageBreakBefore"),
  framePr: readOnce("framePr", (child, formatting) => {
    const frame = parseFrameProperties(child);
    if (frame === undefined) {
      return false;
    }
    formatting.frame = frame;
    return true;
  }),
  widowControl: readToggle("widowControl", "widowControl"),
  numPr: readOnce("numPr", (child, formatting) => {
    const stated = readParagraphNumbering(child);
    if (stated !== undefined) {
      formatting.numPr = stated;
    }
    // These records hold the numbering a reviewer replaced and who
    // inserted the new numbering properties. Nothing derives either from
    // the current model, so a rebuilt `w:numPr` must carry both.
    const numberingChange = findChild(child, "w", "numberingChange");
    if (numberingChange) {
      formatting.numberingChangeXml = captureVerbatimXml(numberingChange);
    }
    const numberingInsertion = findChild(child, "w", "ins");
    if (numberingInsertion) {
      formatting.numberingInsertionXml = captureVerbatimXml(numberingInsertion);
    }
    const statedLevel = parseNumericAttribute(findChild(child, "w", "ilvl"), "w", "val");
    return (
      stated !== undefined ||
      numberingChange !== null ||
      numberingInsertion !== null ||
      statedLevel === -1
    );
  }),
  suppressLineNumbers: readToggle("suppressLineNumbers", "suppressLineNumbers"),
  pBdr: readOnce("pBdr", (child, formatting) => {
    const borders: NonNullable<ParagraphFormatting["borders"]> = {};
    for (const side of PARAGRAPH_BORDER_SIDES) {
      const border = parseBorderSpec(findChild(child, "w", side));
      if (border) {
        borders[side] = border;
      }
    }
    if (Object.keys(borders).length === 0) {
      return false;
    }
    formatting.borders = borders;
    return true;
  }),
  shd: readOnce("shd", (child, formatting) => {
    const shading = parseShading(child);
    if (shading === undefined) {
      return false;
    }
    formatting.shading = shading;
    return true;
  }),
  tabs: readOnce("tabs", (child, formatting) => {
    const tabs = parseTabStops(child);
    if (tabs === undefined) {
      return false;
    }
    formatting.tabs = tabs;
    return true;
  }),
  suppressAutoHyphens: readToggle("suppressAutoHyphens", "suppressAutoHyphens"),
  kinsoku: readToggle("kinsoku", "kinsoku"),
  wordWrap: CAPTURE,
  overflowPunct: readToggle("overflowPunct", "overflowPunctuation"),
  topLinePunct: CAPTURE,
  autoSpaceDE: CAPTURE,
  autoSpaceDN: CAPTURE,
  bidi: readToggle("bidi", "bidi"),
  adjustRightInd: CAPTURE,
  snapToGrid: readToggle("snapToGrid", "snapToGrid"),
  spacing: readOnce("spacing", (child, formatting) => readParagraphSpacing(child, formatting)),
  ind: readOnce("ind", (child, formatting) => readParagraphIndentation(child, formatting)),
  contextualSpacing: readToggle("contextualSpacing", "contextualSpacing"),
  mirrorIndents: CAPTURE,
  suppressOverlap: CAPTURE,
  jc: readOnce("jc", (child, formatting) => {
    const val = narrowEnum(getAttribute(child, "w", "val"), ParagraphAlignmentSchema);
    if (!val) {
      return false;
    }
    formatting.alignment = val;
    return true;
  }),
  textDirection: CAPTURE,
  textAlignment: CAPTURE,
  textboxTightWrap: CAPTURE,
  outlineLvl: readOnce("outlineLvl", (child, formatting) => {
    const val = parseNumericAttribute(child, "w", "val");
    const level = val === undefined ? undefined : outlineLevelFromStatedValue(val);
    if (level === undefined) {
      return val === -1;
    }
    formatting.outlineLevel = level;
    return true;
  }),
  divId: CAPTURE,
  cnfStyle: CAPTURE,
  rPr: ownedElsewhere({
    container: "paragraph-properties",
    child: "rPr",
    reader: "paragraphProperties#parseParagraphProperties",
  }),
  sectPr: ownedElsewhere({
    container: "paragraph-properties",
    child: "sectPr",
    reader: "sectionParser#parseSectionProperties",
  }),
  pPrChange: ownedElsewhere({
    container: "paragraph-properties",
    child: "pPrChange",
    reader: "paragraphParser#parseParagraphPropertyChanges",
  }),
} as const satisfies ChildHandlers<"paragraph-properties", ParagraphPropertyReadContext>;

/**
 * Read `w:pPr` into {@link ParagraphFormatting}, every declared child carrying
 * a decision.
 *
 * One reader serves all four owners of the set — a paragraph's own properties,
 * the `CT_PPrGeneral` a style carries, a numbering level's, and the
 * `CT_PPrBase` snapshot inside `w:pPrChange` — so a property is read the same
 * way wherever it was authored, and the handler map is total over the content
 * model rather than over whichever names a `switch` happened to list.
 *
 * Three children belong to other readers and are neither modelled nor captured
 * here, because capturing one as well would write it twice: `w:sectPr` is read
 * by the section parser off the paragraph, `w:pPrChange` by the
 * property-change parser, and `w:rPr` is the paragraph mark's own property
 * set, read below off the element for the same reason a run's is read off the
 * run.
 */
export function parseParagraphProperties(
  pPr: XmlElement | null,
  theme: Theme | null,
): ParagraphFormatting | undefined {
  if (!pPr) {
    return undefined;
  }

  const formatting: ParagraphFormatting = {};
  const preserved = dispatchChildrenWithContext({
    element: pPr,
    container: "paragraph-properties",
    capturePosition: sequencePositions("paragraph-properties", pPr),
    handlers: PARAGRAPH_PROPERTY_HANDLERS,
    context: { formatting, taken: new Set<string>() },
  });

  // The paragraph mark's own property set. It is `OWNED_ELSEWHERE` above so
  // the sink does not keep a second copy of markup this reader writes back.
  const rPr = findChild(pPr, "w", "rPr");
  if (rPr) {
    const runPropsResult = parseRunProperties(rPr, theme, RUN_PROPERTY_OWNERS.paragraphMark);
    if (runPropsResult !== undefined) {
      formatting.runProperties = runPropsResult;
    }
    // Run-in heading marker: `<w:specVanish/>` on the paragraph
    // mark's rPr (ECMA-376 §17.3.1.32). Word treats the paragraph
    // break as a soft break and flows the next paragraph inline on
    // the same line — used by run-in heading styles in legal
    // templates (NVCA "6.11 Severability" → body merges).
    const specVanish = parseOnOffChild(rPr, "w", "specVanish");
    if (specVanish !== undefined) {
      formatting.runInWithNext = specVanish;
    }
  }

  if (preserved) {
    formatting.preserved = preserved;
  }

  return Object.keys(formatting).length > 0 ? formatting : undefined;
}

/** `w:pBdr`'s sides, in the order `CT_PBdr` declares them. */
const PARAGRAPH_BORDER_SIDES = ["top", "left", "bottom", "right", "between", "bar"] as const;

/**
 * `w:spacing`, which states up to six independent values.
 *
 * @returns whether any of them was taken; an element that states none keeps
 *   its bytes instead, because a `<w:spacing/>` the model cannot rebuild is
 *   markup the author wrote.
 */
const readParagraphSpacing = (spacing: XmlElement, formatting: ParagraphFormatting): boolean => {
  const before = parseNumericAttribute(spacing, "w", "before");
  if (before !== undefined) {
    formatting.spaceBefore = before;
  }

  const after = parseNumericAttribute(spacing, "w", "after");
  if (after !== undefined) {
    formatting.spaceAfter = after;
  }

  const line = parseNumericAttribute(spacing, "w", "line");
  if (line !== undefined) {
    formatting.lineSpacing = line;
  }

  const spacingExplicit: { before?: boolean; after?: boolean } = {};
  if (before !== undefined) {
    spacingExplicit.before = true;
  }
  if (after !== undefined) {
    spacingExplicit.after = true;
  }
  if (spacingExplicit.before || spacingExplicit.after) {
    formatting.spacingExplicit = spacingExplicit;
  }

  const lineRule = narrowEnum(getAttribute(spacing, "w", "lineRule"), LineSpacingRuleSchema);
  if (lineRule) {
    formatting.lineSpacingRule = lineRule;
  }

  const beforeAutospacing = parseOnOffAttribute(spacing, "w", "beforeAutospacing");
  if (beforeAutospacing !== undefined) {
    formatting.beforeAutospacing = beforeAutospacing;
  }

  const afterAutospacing = parseOnOffAttribute(spacing, "w", "afterAutospacing");
  if (afterAutospacing !== undefined) {
    formatting.afterAutospacing = afterAutospacing;
  }

  const taken =
    before !== undefined ||
    after !== undefined ||
    line !== undefined ||
    lineRule !== undefined ||
    beforeAutospacing !== undefined ||
    afterAutospacing !== undefined;
  if (!taken) {
    return false;
  }

  const preservedAttributes = readAttributeBag(spacing, SPACING_ATTRIBUTES);
  if (preservedAttributes) {
    formatting.spacingPreservedAttributes = preservedAttributes;
  }
  return true;
};

/**
 * `w:ind`. `w:start`/`w:end` are the Strict spellings of `w:left`/`w:right`,
 * so they fill the same fields when the Transitional pair is absent.
 *
 * @returns whether any indent was taken.
 */
const readParagraphIndentation = (ind: XmlElement, formatting: ParagraphFormatting): boolean => {
  const left = numericAttributeAnySpelling(ind, "CT_Ind @left");
  if (left !== undefined) {
    formatting.indentLeft = left;
  }

  const right = numericAttributeAnySpelling(ind, "CT_Ind @right");
  if (right !== undefined) {
    formatting.indentRight = right;
  }

  const firstLine = parseNumericAttribute(ind, "w", "firstLine");
  if (firstLine !== undefined) {
    formatting.indentFirstLine = firstLine;
  }

  const hanging = parseNumericAttribute(ind, "w", "hanging");
  if (hanging !== undefined) {
    // Hanging indent is stored as negative first line indent
    formatting.indentFirstLine = -hanging;
    formatting.hangingIndent = true;
  }

  const taken =
    left !== undefined || right !== undefined || firstLine !== undefined || hanging !== undefined;
  if (!taken) {
    return false;
  }

  const preservedAttributes = readAttributeBag(ind, INDENTATION_ATTRIBUTES);
  if (preservedAttributes) {
    formatting.indentPreservedAttributes = preservedAttributes;
  }
  return true;
};
