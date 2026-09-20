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
import { parseBorderSpec } from "./borderParser";
import { CAPTURE, dispatchChildren, OWNED_ELSEWHERE, sequencePositions } from "./containerChildren";
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
import { parseRunProperties } from "./runParser";
import { parseShading } from "./shadingParser";
import type { StyleMap } from "./styleParser";
import { captureVerbatimXml } from "./verbatimCapture";
import {
  findChild,
  findChildren,
  getAttribute,
  parseBooleanElement,
  parseNumericAttribute,
  parseOnOffAttribute,
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

  return Object.keys(frame).length > 0 ? frame : undefined;
}

/** A `w:pPr` child folio models as a tri-state `CT_OnOff` toggle. */
type ParagraphToggleField = {
  [Field in keyof ParagraphFormatting]-?: NonNullable<ParagraphFormatting[Field]> extends boolean
    ? Field
    : never;
}[keyof ParagraphFormatting];

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
  styles?: StyleMap,
): ParagraphFormatting | undefined {
  if (!pPr) {
    return undefined;
  }

  const formatting: ParagraphFormatting = {};
  // The content model declares each property once. A source that states one
  // twice keeps the first as the model's value and the repeat as bytes at the
  // same schema ordinal, rather than letting the second silently win or fall
  // off the end of the walk.
  const taken = new Set<string>();
  const once =
    (name: string, read: (child: XmlElement) => boolean) =>
    (child: XmlElement): typeof CAPTURE | undefined => {
      if (taken.has(name) || !read(child)) {
        return CAPTURE;
      }
      taken.add(name);
      return undefined;
    };
  const toggle = (name: string, field: ParagraphToggleField) =>
    once(name, (child) => {
      formatting[field] = parseBooleanElement(child);
      return true;
    });

  const preserved = dispatchChildren({
    element: pPr,
    container: "paragraph-properties",
    capturePosition: sequencePositions("paragraph-properties", pPr),
    handlers: {
      pStyle: once("pStyle", (child) => {
        const val = getAttribute(child, "w", "val");
        if (!val) {
          return false;
        }
        formatting.styleId = val;
        return true;
      }),
      keepNext: toggle("keepNext", "keepNext"),
      keepLines: toggle("keepLines", "keepLines"),
      pageBreakBefore: toggle("pageBreakBefore", "pageBreakBefore"),
      framePr: once("framePr", (child) => {
        const frame = parseFrameProperties(child);
        if (frame === undefined) {
          return false;
        }
        formatting.frame = frame;
        return true;
      }),
      widowControl: toggle("widowControl", "widowControl"),
      numPr: once("numPr", (child) => {
        const stated = readParagraphNumbering(child);
        if (stated !== undefined) {
          formatting.numPr = stated;
        }
        // `w:numberingChange` records the numbering the paragraph carried
        // before a reviewer changed it. Nothing derives it from the current
        // model, so a rebuilt `w:numPr` that does not carry it discards the
        // revision.
        const numberingChange = findChild(child, "w", "numberingChange");
        if (numberingChange) {
          formatting.numberingChangeXml = captureVerbatimXml(numberingChange);
        }
        return stated !== undefined || numberingChange !== null;
      }),
      suppressLineNumbers: toggle("suppressLineNumbers", "suppressLineNumbers"),
      pBdr: once("pBdr", (child) => {
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
      shd: once("shd", (child) => {
        const shading = parseShading(child);
        if (shading === undefined) {
          return false;
        }
        formatting.shading = shading;
        return true;
      }),
      tabs: once("tabs", (child) => {
        const tabs = parseTabStops(child);
        if (tabs === undefined) {
          return false;
        }
        formatting.tabs = tabs;
        return true;
      }),
      suppressAutoHyphens: toggle("suppressAutoHyphens", "suppressAutoHyphens"),
      kinsoku: toggle("kinsoku", "kinsoku"),
      wordWrap: CAPTURE,
      overflowPunct: toggle("overflowPunct", "overflowPunctuation"),
      topLinePunct: CAPTURE,
      autoSpaceDE: CAPTURE,
      autoSpaceDN: CAPTURE,
      bidi: toggle("bidi", "bidi"),
      adjustRightInd: CAPTURE,
      snapToGrid: toggle("snapToGrid", "snapToGrid"),
      spacing: once("spacing", (child) => readParagraphSpacing(child, formatting)),
      ind: once("ind", (child) => readParagraphIndentation(child, formatting)),
      contextualSpacing: toggle("contextualSpacing", "contextualSpacing"),
      mirrorIndents: CAPTURE,
      suppressOverlap: CAPTURE,
      jc: once("jc", (child) => {
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
      outlineLvl: once("outlineLvl", (child) => {
        const val = parseNumericAttribute(child, "w", "val");
        const level = val === undefined ? undefined : outlineLevelFromStatedValue(val);
        if (level === undefined) {
          return false;
        }
        formatting.outlineLevel = level;
        return true;
      }),
      divId: CAPTURE,
      cnfStyle: CAPTURE,
      rPr: OWNED_ELSEWHERE,
      sectPr: OWNED_ELSEWHERE,
      pPrChange: OWNED_ELSEWHERE,
    },
  });

  // The paragraph mark's own property set. It is `OWNED_ELSEWHERE` above so
  // the sink does not keep a second copy of markup this reader writes back.
  const rPr = findChild(pPr, "w", "rPr");
  if (rPr) {
    const runPropsResult = parseRunProperties(rPr, theme, styles);
    if (runPropsResult !== undefined) {
      formatting.runProperties = runPropsResult;
    }
    // Run-in heading marker: `<w:specVanish/>` on the paragraph
    // mark's rPr (ECMA-376 §17.3.1.32). Word treats the paragraph
    // break as a soft break and flows the next paragraph inline on
    // the same line — used by run-in heading styles in legal
    // templates (NVCA "6.11 Severability" → body merges).
    const specVanish = findChild(rPr, "w", "specVanish");
    if (specVanish && parseBooleanElement(specVanish)) {
      formatting.runInWithNext = true;
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

  return (
    before !== undefined ||
    after !== undefined ||
    line !== undefined ||
    lineRule !== undefined ||
    beforeAutospacing !== undefined ||
    afterAutospacing !== undefined
  );
};

/**
 * `w:ind`. `w:start`/`w:end` are the Strict spellings of `w:left`/`w:right`,
 * so they fill the same fields when the Transitional pair is absent.
 *
 * @returns whether any indent was taken.
 */
const readParagraphIndentation = (ind: XmlElement, formatting: ParagraphFormatting): boolean => {
  const left = parseNumericAttribute(ind, "w", "left");
  if (left !== undefined) {
    formatting.indentLeft = left;
  }

  const right = parseNumericAttribute(ind, "w", "right");
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

  const start = parseNumericAttribute(ind, "w", "start");
  if (start !== undefined && formatting.indentLeft === undefined) {
    formatting.indentLeft = start;
  }

  const end = parseNumericAttribute(ind, "w", "end");
  if (end !== undefined && formatting.indentRight === undefined) {
    formatting.indentRight = end;
  }

  return (
    left !== undefined ||
    right !== undefined ||
    firstLine !== undefined ||
    hanging !== undefined ||
    start !== undefined ||
    end !== undefined
  );
};
