import type { ExhaustiveFields, ParagraphFormatting, TextFormatting } from "../types/document";
import { serializeBorder } from "../docx/serializer/borderSerializer";
import {
  serializeShading,
  serializeTextFormatting,
} from "../docx/serializer/textFormattingSerializer";
import { intAttr } from "../docx/serializer/xmlUtils";
import { escapeXmlAttribute } from "@stll/docx-core";
import {
  outlineLevelStatedValue,
  paragraphNumberingSlots,
  sameEffectiveParagraphNumbering,
  type ParagraphNumberingOverride,
} from "@stll/docx-core/model";
import { TRANSITIONAL_NAME_BY_STRICT_NAME } from "../docx/strictNames.gen";
import { sanitizeCapturedXmlElement } from "../docx/verbatimCapture";
import { NAMESPACES, OOXML_NAMESPACE_SCOPE } from "../docx/xmlParser";

type RequiredFieldValues<Source, Fields extends keyof Source> = {
  [Field in Fields]: Source[Field] | undefined;
};

type ParagraphBorders = NonNullable<ParagraphFormatting["borders"]>;
type ClassifiedParagraphBordersField = "top" | "left" | "bottom" | "right" | "between" | "bar";
type ExhaustiveParagraphBorders = ExhaustiveFields<
  ParagraphBorders,
  ClassifiedParagraphBordersField
>;

type ParagraphTab = NonNullable<ParagraphFormatting["tabs"]>[number];
type ClassifiedParagraphTabField = "alignment" | "position" | "leader";
type ExhaustiveParagraphTab = ExhaustiveFields<ParagraphTab, ClassifiedParagraphTabField>;

type ParagraphFrame = NonNullable<ParagraphFormatting["frame"]>;
type ClassifiedParagraphFrameField =
  | "dropCap"
  | "lines"
  | "width"
  | "height"
  | "hSpace"
  | "vSpace"
  | "hAnchor"
  | "vAnchor"
  | "x"
  | "y"
  | "xAlign"
  | "yAlign"
  | "wrap";
type ExhaustiveParagraphFrame = ExhaustiveFields<ParagraphFrame, ClassifiedParagraphFrameField>;

type SpacingProvenance = NonNullable<ParagraphFormatting["spacingExplicit"]>;
type ClassifiedSpacingProvenanceField = "before" | "after";
type ExhaustiveSpacingProvenance = ExhaustiveFields<
  SpacingProvenance,
  ClassifiedSpacingProvenanceField
>;

type ClassifiedParagraphFormattingField =
  | "alignment"
  | "bidi"
  | "kinsoku"
  | "overflowPunctuation"
  | "spaceBefore"
  | "spaceAfter"
  | "lineSpacing"
  | "lineSpacingRule"
  | "snapToGrid"
  | "beforeAutospacing"
  | "afterAutospacing"
  | "spacingExplicit"
  | "indentLeft"
  | "indentRight"
  | "indentFirstLine"
  | "hangingIndent"
  | "borders"
  | "shading"
  | "tabs"
  | "keepNext"
  | "keepLines"
  | "widowControl"
  | "pageBreakBefore"
  | "contextualSpacing"
  | "numPr"
  | "numPrFromStyle"
  | "numberingChangeXml"
  | "outlineLevel"
  | "styleId"
  | "frame"
  | "suppressLineNumbers"
  | "suppressAutoHyphens"
  | "runProperties"
  | "runInWithNext";
type ExhaustiveParagraphFormatting = ExhaustiveFields<
  ParagraphFormatting,
  ClassifiedParagraphFormattingField
>;

/**
 * Whether resolved numbering still belongs to the paragraph's style tier.
 *
 * Effective equality, not stated: the resolved field carries the level the
 * cascade supplied even where the style stated none, and a paragraph whose
 * numbering came wholly from its style must still not emit a direct
 * `<w:numPr>`.
 */
export const isStyleSourcedParagraphNumbering = (
  numPr: ParagraphNumberingOverride | null | undefined,
  numPrFromStyle: ParagraphNumberingOverride | null | undefined,
): boolean =>
  numPr != null && numPrFromStyle != null && sameEffectiveParagraphNumbering(numPr, numPrFromStyle);

/** Exact fallback-emission instructions for the modeled part of `w:pPr`. */
export type ModeledParagraphFormattingEmission = Readonly<{
  propertiesXml?: string;
  paragraphMarkPropertiesInnerXml?: string;
}>;

type MutableModeledParagraphFormattingEmission = {
  -readonly [Field in keyof ModeledParagraphFormattingEmission]: ModeledParagraphFormattingEmission[Field];
};

const serializeToggle = (name: string, value: boolean | undefined): string => {
  if (value === true) {
    return `<w:${name}/>`;
  }
  if (value === false) {
    return `<w:${name} w:val="0"/>`;
  }
  return "";
};

const serializeParagraphBorders = (borders: ExhaustiveParagraphBorders | undefined): string => {
  if (!borders) {
    return "";
  }

  const { top, left, bottom, right, between, bar } = borders;

  const parts = [
    serializeBorder(top, "top"),
    serializeBorder(left, "left"),
    serializeBorder(bottom, "bottom"),
    serializeBorder(right, "right"),
    serializeBorder(between, "between"),
    serializeBorder(bar, "bar"),
  ].filter(Boolean);
  return parts.length === 0 ? "" : `<w:pBdr>${parts.join("")}</w:pBdr>`;
};

const serializeTabStops = (tabs: ParagraphFormatting["tabs"]): string => {
  if (!tabs || tabs.length === 0) {
    return "";
  }

  const tabElements = tabs.map((tab) => {
    const exhaustiveTab: ExhaustiveParagraphTab = tab;
    const { alignment, position, leader } = exhaustiveTab;
    const attrs = [`w:val="${alignment}"`, `w:pos="${intAttr(position)}"`];
    if (leader && leader !== "none") {
      attrs.push(`w:leader="${leader}"`);
    }
    return `<w:tab ${attrs.join(" ")}/>`;
  });
  return `<w:tabs>${tabElements.join("")}</w:tabs>`;
};

type ClassifiedSpacingFormattingField =
  | "spaceBefore"
  | "spaceAfter"
  | "lineSpacing"
  | "lineSpacingRule"
  | "beforeAutospacing"
  | "afterAutospacing";
type SpacingFormatting = RequiredFieldValues<ParagraphFormatting, ClassifiedSpacingFormattingField>;

const serializeSpacing = (formatting: SpacingFormatting): string => {
  const attrs: string[] = [];
  if (formatting.spaceBefore !== undefined) {
    attrs.push(`w:before="${intAttr(formatting.spaceBefore)}"`);
  }
  if (formatting.spaceAfter !== undefined) {
    attrs.push(`w:after="${intAttr(formatting.spaceAfter)}"`);
  }
  if (formatting.lineSpacing !== undefined) {
    attrs.push(`w:line="${intAttr(formatting.lineSpacing)}"`);
  }
  if (formatting.lineSpacingRule) {
    attrs.push(`w:lineRule="${formatting.lineSpacingRule}"`);
  }
  if (formatting.beforeAutospacing !== undefined) {
    attrs.push(`w:beforeAutospacing="${formatting.beforeAutospacing ? "1" : "0"}"`);
  }
  if (formatting.afterAutospacing !== undefined) {
    attrs.push(`w:afterAutospacing="${formatting.afterAutospacing ? "1" : "0"}"`);
  }
  return attrs.length === 0 ? "" : `<w:spacing ${attrs.join(" ")}/>`;
};

type ClassifiedIndentationFormattingField =
  | "indentLeft"
  | "indentRight"
  | "indentFirstLine"
  | "hangingIndent";
type IndentationFormatting = RequiredFieldValues<
  ParagraphFormatting,
  ClassifiedIndentationFormattingField
>;

/**
 * The edge names folio writes, from the table the census reads.
 *
 * A Strict producer spells these `@w:start` and `@w:end`; folio writes one
 * spelling, and taking it from the generated table is what keeps the writer and
 * the survival law's equivalence from drifting apart.
 */
const INDENT_LEFT = TRANSITIONAL_NAME_BY_STRICT_NAME["CT_Ind @start"];
const INDENT_RIGHT = TRANSITIONAL_NAME_BY_STRICT_NAME["CT_Ind @end"];

const serializeIndentation = (formatting: IndentationFormatting): string => {
  const attrs: string[] = [];
  if (formatting.indentLeft !== undefined) {
    attrs.push(`w:${INDENT_LEFT}="${intAttr(formatting.indentLeft)}"`);
  }
  if (formatting.indentRight !== undefined) {
    attrs.push(`w:${INDENT_RIGHT}="${intAttr(formatting.indentRight)}"`);
  }
  if (formatting.indentFirstLine !== undefined) {
    const attribute = formatting.hangingIndent ? "hanging" : "firstLine";
    const value = formatting.hangingIndent
      ? Math.abs(formatting.indentFirstLine)
      : formatting.indentFirstLine;
    attrs.push(`w:${attribute}="${intAttr(value)}"`);
  }
  return attrs.length === 0 ? "" : `<w:ind ${attrs.join(" ")}/>`;
};

/**
 * `w:numPr`, with the tracked record of the numbering it replaced.
 *
 * `w:numberingChange` is the revision a reviewer's numbering change left
 * behind. It is history: nothing in the model derives it, and a rebuilt
 * `w:numPr` that drops it discards the revision silently. It is written even
 * when the paragraph's own numbering reference is gone, because a change
 * record with nothing left to describe is still a record — the schema declares
 * it last in `CT_NumPr`, after `w:ilvl` and `w:numId`.
 */
const serializeNumbering = (
  numPr: ParagraphFormatting["numPr"],
  numberingChangeXml: string | undefined,
): string => {
  const slots = numPr === undefined ? {} : paragraphNumberingSlots(numPr);
  const parts: string[] = [];
  if (slots.ilvl !== undefined) {
    parts.push(`<w:ilvl w:val="${intAttr(slots.ilvl)}"/>`);
  }
  if (slots.numId !== undefined) {
    parts.push(`<w:numId w:val="${intAttr(slots.numId)}"/>`);
  }
  const change = replayableNumberingChangeXml(numberingChangeXml);
  if (change !== null) {
    parts.push(change);
  }
  return parts.length === 0 ? "" : `<w:numPr>${parts.join("")}</w:numPr>`;
};

const NUMBERING_CHANGE_ROOT_NAME: ReadonlySet<string> = new Set(["numberingChange"]);
const WORDPROCESSINGML_NAMESPACE: ReadonlySet<string> = new Set([NAMESPACES.w]);

const replayableNumberingChangeXml = (numberingChangeXml: string | undefined): string | null =>
  sanitizeCapturedXmlElement(numberingChangeXml, {
    allowedLocalNames: NUMBERING_CHANGE_ROOT_NAME,
    allowedNamespaceUris: WORDPROCESSINGML_NAMESPACE,
    inheritedNamespaceScope: OOXML_NAMESPACE_SCOPE,
  });

const serializeFrameProperties = (frame: ParagraphFormatting["frame"]): string => {
  if (!frame) {
    return "";
  }

  const exhaustiveFrame: ExhaustiveParagraphFrame = frame;
  const {
    dropCap,
    lines,
    width,
    height,
    hSpace,
    vSpace,
    hAnchor,
    vAnchor,
    x,
    y,
    xAlign,
    yAlign,
    wrap,
  } = exhaustiveFrame;

  const attrs: string[] = [];
  if (dropCap) attrs.push(`w:dropCap="${dropCap}"`);
  if (lines !== undefined) attrs.push(`w:lines="${intAttr(lines)}"`);
  if (width !== undefined) attrs.push(`w:w="${intAttr(width)}"`);
  if (height !== undefined) attrs.push(`w:h="${intAttr(height)}"`);
  if (hSpace !== undefined) attrs.push(`w:hSpace="${intAttr(hSpace)}"`);
  if (vSpace !== undefined) attrs.push(`w:vSpace="${intAttr(vSpace)}"`);
  if (hAnchor) attrs.push(`w:hAnchor="${hAnchor}"`);
  if (vAnchor) attrs.push(`w:vAnchor="${vAnchor}"`);
  if (x !== undefined) attrs.push(`w:x="${x}"`);
  if (y !== undefined) attrs.push(`w:y="${y}"`);
  if (xAlign) attrs.push(`w:xAlign="${xAlign}"`);
  if (yAlign) attrs.push(`w:yAlign="${yAlign}"`);
  if (wrap) attrs.push(`w:wrap="${wrap}"`);
  return attrs.length === 0 ? "" : `<w:framePr ${attrs.join(" ")}/>`;
};

const modelSpacingProvenance = (
  spacingExplicit: ParagraphFormatting["spacingExplicit"],
): undefined => {
  if (spacingExplicit) {
    const exhaustiveProvenance: ExhaustiveSpacingProvenance = spacingExplicit;
    const { before: _before, after: _after } = exhaustiveProvenance;
  }
  return undefined;
};

const MARK_PROPERTIES_OPEN = "<w:rPr>";
const MARK_PROPERTIES_CLOSE = "</w:rPr>";
const MARK_PROPERTIES_EMPTY = "<w:rPr/>";

/**
 * The paragraph mark's `w:rPr`, without its element, as the one `w:rPr` writer
 * composes it.
 *
 * Three answers rather than two: `undefined` when the mark carried no property
 * set, `""` when it carried an empty one, and the inner markup otherwise.
 * Collapsing the middle case onto the first is what dropped `<w:rPr/>`: the
 * element is optional, so writing one is not the same as writing none, and on
 * the paragraph mark it is the slot a revision on the mark lives in.
 *
 * The caller re-wraps it with the mark's own revision in front, so this side
 * hands back the inner markup rather than the element. `w:specVanish` travels
 * as an owned child instead of being appended, because `EG_RPrBase` declares
 * it between `w:eastAsianLayout` and `w:oMath` rather than last.
 */
const paragraphMarkPropertiesInner = (
  runProperties: TextFormatting | undefined,
  runInWithNext: boolean | undefined,
): string | undefined => {
  const xml = serializeTextFormatting(
    runProperties,
    runInWithNext === true ? [["specVanish", "<w:specVanish/>"]] : [],
  );
  // One branch per answer the writer has: nothing, the empty element, the
  // element with children. A fourth branch would be guessing at a shape the
  // writer cannot produce.
  if (xml === "") {
    return undefined;
  }
  if (xml === MARK_PROPERTIES_EMPTY) {
    return "";
  }
  return xml.slice(MARK_PROPERTIES_OPEN.length, -MARK_PROPERTIES_CLOSE.length);
};

/**
 * Resolve paragraph formatting into its exact fallback `w:pPr` instructions.
 * Every input field is classified here so parser capture fingerprints and the
 * serializer cannot drift apart when paragraph formatting grows.
 */
export const modelParagraphFormattingEmission = (
  input: ExhaustiveParagraphFormatting | undefined,
): ModeledParagraphFormattingEmission => {
  if (!input) {
    return {};
  }
  const formatting: ExhaustiveParagraphFormatting = input;
  const {
    alignment,
    bidi,
    kinsoku,
    overflowPunctuation,
    spaceBefore,
    spaceAfter,
    lineSpacing,
    lineSpacingRule,
    snapToGrid,
    beforeAutospacing,
    afterAutospacing,
    spacingExplicit,
    indentLeft,
    indentRight,
    indentFirstLine,
    hangingIndent,
    borders,
    shading,
    tabs,
    keepNext,
    keepLines,
    widowControl,
    pageBreakBefore,
    contextualSpacing,
    numPr,
    numPrFromStyle,
    numberingChangeXml,
    outlineLevel,
    styleId,
    frame,
    suppressLineNumbers,
    suppressAutoHyphens,
    runProperties,
    runInWithNext,
  } = formatting;

  modelSpacingProvenance(spacingExplicit);

  const properties = [
    styleId ? `<w:pStyle w:val="${escapeXmlAttribute(styleId)}"/>` : "",
    serializeToggle("keepNext", keepNext),
    serializeToggle("keepLines", keepLines),
    serializeToggle("pageBreakBefore", pageBreakBefore),
    serializeFrameProperties(frame),
    serializeToggle("widowControl", widowControl),
    isStyleSourcedParagraphNumbering(numPr, numPrFromStyle)
      ? serializeNumbering(undefined, numberingChangeXml)
      : serializeNumbering(numPr, numberingChangeXml),
    serializeToggle("suppressLineNumbers", suppressLineNumbers),
    serializeParagraphBorders(borders),
    serializeShading(shading),
    serializeTabStops(tabs),
    serializeToggle("suppressAutoHyphens", suppressAutoHyphens),
    serializeToggle("kinsoku", kinsoku),
    serializeToggle("overflowPunct", overflowPunctuation),
    serializeToggle("bidi", bidi),
    serializeToggle("snapToGrid", snapToGrid),
    serializeSpacing({
      spaceBefore,
      spaceAfter,
      lineSpacing,
      lineSpacingRule,
      beforeAutospacing,
      afterAutospacing,
    }),
    serializeIndentation({ indentLeft, indentRight, indentFirstLine, hangingIndent }),
    serializeToggle("contextualSpacing", contextualSpacing),
    alignment ? `<w:jc w:val="${alignment}"/>` : "",
    outlineLevel === undefined
      ? ""
      : `<w:outlineLvl w:val="${outlineLevelStatedValue(outlineLevel)}"/>`,
  ];
  const propertiesXml = properties.join("");
  const paragraphMarkPropertiesInnerXml = paragraphMarkPropertiesInner(
    runProperties,
    runInWithNext,
  );
  const emission: MutableModeledParagraphFormattingEmission = {};
  if (propertiesXml) emission.propertiesXml = propertiesXml;
  if (paragraphMarkPropertiesInnerXml !== undefined) {
    emission.paragraphMarkPropertiesInnerXml = paragraphMarkPropertiesInnerXml;
  }

  return emission;
};
