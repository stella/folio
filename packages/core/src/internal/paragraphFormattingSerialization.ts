import type { ParagraphFormatting } from "../types/document";
import { serializeBorder } from "../docx/serializer/borderSerializer";
import {
  serializeShading,
  serializeTextFormatting,
} from "../docx/serializer/textFormattingSerializer";
import { escapeXml, intAttr } from "../docx/serializer/xmlUtils";

type ExhaustiveFields<Source, Classified extends keyof Source> =
  Exclude<keyof Source, Classified> extends never ? Source : never;

type RequiredFieldValues<Source, Fields extends keyof Source> = {
  [Field in Fields]: Source[Field] | undefined;
};

type ParagraphNumberingReference = ParagraphFormatting["numPr"] | null;
type ParagraphNumbering = NonNullable<ParagraphFormatting["numPr"]>;
type StyleParagraphNumbering = NonNullable<ParagraphFormatting["numPrFromStyle"]>;
type ClassifiedParagraphNumberingField = "numId" | "ilvl";
type ExhaustiveParagraphNumbering = ExhaustiveFields<
  ParagraphNumbering,
  ClassifiedParagraphNumberingField
> &
  ExhaustiveFields<StyleParagraphNumbering, ClassifiedParagraphNumberingField>;

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

type ModeledParagraphNumberingReference = Readonly<{
  numId?: number;
  ilvl?: number;
}>;

const modelParagraphNumberingReference = (
  reference: ExhaustiveParagraphNumbering | null | undefined,
): ModeledParagraphNumberingReference | null => {
  if (!reference) {
    return null;
  }
  const { numId, ilvl } = reference;
  const modeled = {
    ...(numId !== undefined ? { numId } : {}),
    ...(ilvl !== undefined ? { ilvl } : {}),
  };
  return modeled;
};

/** Compare numbering references by their emitted id and effective level. */
export const paragraphNumberingReferencesEqual = (
  left: ParagraphNumberingReference,
  right: ParagraphNumberingReference,
): boolean => {
  const modeledLeft = modelParagraphNumberingReference(left);
  const modeledRight = modelParagraphNumberingReference(right);
  if (modeledLeft === null || modeledRight === null) {
    return modeledLeft === null && modeledRight === null;
  }
  return (
    modeledLeft.numId === modeledRight.numId && (modeledLeft.ilvl ?? 0) === (modeledRight.ilvl ?? 0)
  );
};

/** Whether resolved numbering still belongs to the paragraph's style tier. */
export const isStyleSourcedParagraphNumbering = (
  numPr: ParagraphNumberingReference,
  numPrFromStyle: ParagraphNumberingReference,
): boolean =>
  numPr != null &&
  numPrFromStyle != null &&
  paragraphNumberingReferencesEqual(numPr, numPrFromStyle);

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

const serializeIndentation = (formatting: IndentationFormatting): string => {
  const attrs: string[] = [];
  if (formatting.indentLeft !== undefined) {
    attrs.push(`w:left="${intAttr(formatting.indentLeft)}"`);
  }
  if (formatting.indentRight !== undefined) {
    attrs.push(`w:right="${intAttr(formatting.indentRight)}"`);
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

const serializeNumbering = (numPr: ParagraphFormatting["numPr"]): string => {
  const modeled = modelParagraphNumberingReference(numPr ?? null);
  if (!modeled) {
    return "";
  }

  const parts: string[] = [];
  if (modeled.ilvl !== undefined) {
    parts.push(`<w:ilvl w:val="${intAttr(modeled.ilvl)}"/>`);
  }
  if (modeled.numId !== undefined) {
    parts.push(`<w:numId w:val="${intAttr(modeled.numId)}"/>`);
  }
  return parts.length === 0 ? "" : `<w:numPr>${parts.join("")}</w:numPr>`;
};

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

const extractRunPropertiesInnerXml = (runPropertiesXml: string): string => {
  if (!runPropertiesXml.startsWith("<w:rPr>") || !runPropertiesXml.endsWith("</w:rPr>")) {
    return "";
  }
  return runPropertiesXml.slice("<w:rPr>".length, -"</w:rPr>".length);
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
    styleId ? `<w:pStyle w:val="${escapeXml(styleId)}"/>` : "",
    serializeToggle("keepNext", keepNext),
    serializeToggle("keepLines", keepLines),
    serializeToggle("pageBreakBefore", pageBreakBefore),
    serializeFrameProperties(frame),
    serializeToggle("widowControl", widowControl),
    isStyleSourcedParagraphNumbering(numPr, numPrFromStyle) ? "" : serializeNumbering(numPr),
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
    outlineLevel !== undefined ? `<w:outlineLvl w:val="${outlineLevel}"/>` : "",
  ];
  const propertiesXml = properties.join("");
  const runPropertiesInnerXml = extractRunPropertiesInnerXml(
    serializeTextFormatting(runProperties),
  );
  const paragraphMarkPropertiesInnerXml = `${runPropertiesInnerXml}${
    runInWithNext === true ? "<w:specVanish/>" : ""
  }`;
  const emission: MutableModeledParagraphFormattingEmission = {};
  if (propertiesXml) emission.propertiesXml = propertiesXml;
  if (paragraphMarkPropertiesInnerXml) {
    emission.paragraphMarkPropertiesInnerXml = paragraphMarkPropertiesInnerXml;
  }

  return emission;
};
