import type { ExhaustiveFields, ParagraphFormatting, TextFormatting } from "../types/document";
import { serializePreservedAttributes } from "../docx/attributeRemainder";
import { serializeBorder } from "../docx/serializer/borderSerializer";
import {
  serializeShading,
  serializeTextFormatting,
} from "../docx/serializer/textFormattingSerializer";
import { intAttr } from "../docx/serializer/xmlUtils";
import { escapeXmlAttribute, serializeOnOffElement } from "@stll/docx-core";
import {
  outlineLevelStatedValue,
  paragraphNumberingSlots,
  sameEffectiveParagraphNumbering,
  type ParagraphNumberingOverride,
} from "@stll/docx-core/model";
import { serializeSequenceChildren } from "@stll/docx-core/schema";
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
type ClassifiedParagraphTabField = "alignment" | "position" | "leader" | "preservedAttributes";
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
  | "wrap"
  | "preservedAttributes";
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
  | "spacingPreservedAttributes"
  | "indentLeft"
  | "indentRight"
  | "indentFirstLine"
  | "hangingIndent"
  | "indentPreservedAttributes"
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
  | "numberingInsertionXml"
  | "outlineLevel"
  | "styleId"
  | "frame"
  | "suppressLineNumbers"
  | "suppressAutoHyphens"
  | "runProperties"
  | "runInWithNext"
  | "preserved";
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
    const { alignment, position, leader, preservedAttributes } = exhaustiveTab;
    const attrs = [`w:val="${alignment}"`, `w:pos="${intAttr(position)}"`];
    if (leader && leader !== "none") {
      attrs.push(`w:leader="${leader}"`);
    }
    return `<w:tab ${serializePreservedAttributes(attrs, preservedAttributes).join(" ")}/>`;
  });
  return `<w:tabs>${tabElements.join("")}</w:tabs>`;
};

type ClassifiedSpacingFormattingField =
  | "spaceBefore"
  | "spaceAfter"
  | "lineSpacing"
  | "lineSpacingRule"
  | "beforeAutospacing"
  | "afterAutospacing"
  | "spacingPreservedAttributes";
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
  const written = serializePreservedAttributes(attrs, formatting.spacingPreservedAttributes);
  return written.length === 0 ? "" : `<w:spacing ${written.join(" ")}/>`;
};

type ClassifiedIndentationFormattingField =
  | "indentLeft"
  | "indentRight"
  | "indentFirstLine"
  | "hangingIndent"
  | "indentPreservedAttributes";
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
  const written = serializePreservedAttributes(attrs, formatting.indentPreservedAttributes);
  return written.length === 0 ? "" : `<w:ind ${written.join(" ")}/>`;
};

/**
 * `w:numPr`, with the tracked records attached to its numbering properties.
 *
 * `w:numberingChange` is the revision a reviewer's numbering change left
 * behind; `w:ins` records who inserted the numbering properties. Neither is
 * derived by the model, and dropping either silently changes the review
 * history. `CT_NumPr` declares both after `w:ilvl` and `w:numId`, with
 * `w:numberingChange` before `w:ins`.
 */
type SerializeNumberingOptions = {
  numPr: ParagraphFormatting["numPr"];
  numberingChangeXml: string | undefined;
  numberingInsertionXml: string | undefined;
};

const serializeNumbering = ({
  numPr,
  numberingChangeXml,
  numberingInsertionXml,
}: SerializeNumberingOptions): string => {
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
  const insertion = replayableNumberingInsertionXml(numberingInsertionXml);
  if (insertion !== null) {
    parts.push(insertion);
  }
  return parts.length === 0 ? "" : `<w:numPr>${parts.join("")}</w:numPr>`;
};

const NUMBERING_CHANGE_ROOT_NAME: ReadonlySet<string> = new Set(["numberingChange"]);
const NUMBERING_INSERTION_ROOT_NAME: ReadonlySet<string> = new Set(["ins"]);
const WORDPROCESSINGML_NAMESPACE: ReadonlySet<string> = new Set([NAMESPACES.w]);

const replayableNumberingChangeXml = (numberingChangeXml: string | undefined): string | null =>
  sanitizeCapturedXmlElement(numberingChangeXml, {
    allowedLocalNames: NUMBERING_CHANGE_ROOT_NAME,
    allowedNamespaceUris: WORDPROCESSINGML_NAMESPACE,
    inheritedNamespaceScope: OOXML_NAMESPACE_SCOPE,
  });

const replayableNumberingInsertionXml = (
  numberingInsertionXml: string | undefined,
): string | null =>
  sanitizeCapturedXmlElement(numberingInsertionXml, {
    allowedLocalNames: NUMBERING_INSERTION_ROOT_NAME,
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
    preservedAttributes,
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
  const written = serializePreservedAttributes(attrs, preservedAttributes);
  return written.length === 0 ? "" : `<w:framePr ${written.join(" ")}/>`;
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
    runInWithNext === undefined
      ? []
      : [["specVanish", serializeOnOffElement(runInWithNext, "specVanish")]],
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
    spacingPreservedAttributes,
    indentLeft,
    indentRight,
    indentFirstLine,
    hangingIndent,
    indentPreservedAttributes,
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
    numberingInsertionXml,
    outlineLevel,
    styleId,
    frame,
    suppressLineNumbers,
    suppressAutoHyphens,
    runProperties,
    runInWithNext,
    preserved,
  } = formatting;

  modelSpacingProvenance(spacingExplicit);

  const propertiesXml = serializeSequenceChildren({
    container: "paragraph-properties",
    preserved,
    modelled: [
      ["pStyle", styleId ? `<w:pStyle w:val="${escapeXmlAttribute(styleId)}"/>` : ""],
      ["keepNext", serializeOnOffElement(keepNext, "keepNext")],
      ["keepLines", serializeOnOffElement(keepLines, "keepLines")],
      ["pageBreakBefore", serializeOnOffElement(pageBreakBefore, "pageBreakBefore")],
      ["framePr", serializeFrameProperties(frame)],
      ["widowControl", serializeOnOffElement(widowControl, "widowControl")],
      [
        "numPr",
        isStyleSourcedParagraphNumbering(numPr, numPrFromStyle)
          ? serializeNumbering({
              numPr: undefined,
              numberingChangeXml,
              numberingInsertionXml,
            })
          : serializeNumbering({ numPr, numberingChangeXml, numberingInsertionXml }),
      ],
      ["suppressLineNumbers", serializeOnOffElement(suppressLineNumbers, "suppressLineNumbers")],
      ["pBdr", serializeParagraphBorders(borders)],
      ["shd", serializeShading(shading)],
      ["tabs", serializeTabStops(tabs)],
      ["suppressAutoHyphens", serializeOnOffElement(suppressAutoHyphens, "suppressAutoHyphens")],
      ["kinsoku", serializeOnOffElement(kinsoku, "kinsoku")],
      ["overflowPunct", serializeOnOffElement(overflowPunctuation, "overflowPunct")],
      ["bidi", serializeOnOffElement(bidi, "bidi")],
      ["snapToGrid", serializeOnOffElement(snapToGrid, "snapToGrid")],
      [
        "spacing",
        serializeSpacing({
          spaceBefore,
          spaceAfter,
          lineSpacing,
          lineSpacingRule,
          beforeAutospacing,
          afterAutospacing,
          spacingPreservedAttributes,
        }),
      ],
      [
        "ind",
        serializeIndentation({
          indentLeft,
          indentRight,
          indentFirstLine,
          hangingIndent,
          indentPreservedAttributes,
        }),
      ],
      ["contextualSpacing", serializeOnOffElement(contextualSpacing, "contextualSpacing")],
      ["jc", alignment ? `<w:jc w:val="${alignment}"/>` : ""],
      [
        "outlineLvl",
        outlineLevel === undefined
          ? ""
          : `<w:outlineLvl w:val="${outlineLevelStatedValue(outlineLevel)}"/>`,
      ],
    ],
  }).join("");
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

type ParagraphPropertySetOptions = {
  formatting: ParagraphFormatting | undefined;
  markPropertiesPrefixXml?: string;
  sectionPropertiesXml?: string;
  propertyChangesXml?: readonly string[];
};

/** Serialize the shared `w:pPr` shape in schema order for all four owners. */
export const serializeParagraphPropertySet = ({
  formatting,
  markPropertiesPrefixXml = "",
  sectionPropertiesXml = "",
  propertyChangesXml = [],
}: ParagraphPropertySetOptions): string => {
  const modeled = modelParagraphFormattingEmission(formatting);
  const markProperties = modeled.paragraphMarkPropertiesInnerXml;
  const markInner = `${markPropertiesPrefixXml}${markProperties ?? ""}`;
  let markXml = "";
  if (markPropertiesPrefixXml !== "" || markProperties !== undefined) {
    markXml = markInner === "" ? "<w:rPr/>" : `<w:rPr>${markInner}</w:rPr>`;
  }
  const inner = [
    modeled.propertiesXml ?? "",
    markXml,
    sectionPropertiesXml,
    ...propertyChangesXml,
  ].join("");
  return inner === "" ? "" : `<w:pPr>${inner}</w:pPr>`;
};
