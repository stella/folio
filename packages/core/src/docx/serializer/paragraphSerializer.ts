/**
 * Paragraph Serializer - Serialize paragraphs to OOXML XML
 *
 * Converts Paragraph objects back to <w:p> XML format for DOCX files.
 * Handles all paragraph properties and child content (runs, hyperlinks, fields, bookmarks).
 *
 * OOXML Reference:
 * - Paragraph: w:p
 * - Paragraph properties: w:pPr
 * - Runs, hyperlinks, bookmarks, fields as child elements
 */

import type {
  Paragraph,
  ParagraphContent,
  ParagraphFormatting,
  ParagraphMarkChange,
  Run,
  Hyperlink,
  BookmarkStart,
  BookmarkEnd,
  SimpleField,
  ComplexField,
  InlineSdt,
  Insertion,
  Deletion,
  MoveFrom,
  MoveTo,
  MoveFromRangeStart,
  MoveToRangeStart,
  ParagraphPropertyChange,
  SectionProperties,
  SdtProperties,
  TabStop,
  ShadingProperties,
  TextFormatting,
} from "../../types/document";
import {
  canonicalParagraphPropertySourceFingerprintJson,
  PARAGRAPH_MARK_CHANGE_KINDS,
  paragraphPropertySourceFingerprintFromFormatting,
} from "@stll/docx-core/model";
import { panic } from "better-result";
import { isValidHexColor } from "../../utils/colorResolver";
import { numPrEqual } from "../numberingParser";
import { getParagraphPropertySource } from "../paragraphPropertySource";
import { reconcileRawSdtPr } from "../sdtPropertiesPatch";
import { DATE_UTC_ATTRIBUTE, DATE_UTC_NAMESPACE_URI } from "../trackedChangeInfo";
import { toTransitionalNamespaceUri } from "../transitionalSpelling";
import { captureVerbatimXml, sanitizeCapturedXmlElement } from "../verbatimCapture";
import {
  cloneElement,
  getChildElements,
  getLocalName,
  NAMESPACES,
  OOXML_NAMESPACE_SCOPE,
  parseXml,
  type XmlElement,
  type XmlNamespaceScope,
} from "../xmlParser";
import { serializeBorder } from "./borderSerializer";
// oxlint-disable-next-line import/no-cycle -- OOXML model is mutually recursive: paragraphs hold runs, shape-textbox runs hold paragraphs
import { serializeRun, serializeTextFormatting } from "./runSerializer";
import { serializeSectionProperties } from "./sectionPropertiesSerializer";
import {
  serializeTrackedChangeAttributes,
  trackedChangeAttributeRecord,
} from "./trackedChangeAttributes";
import { escapeXml, intAttr, isSingleWellFormedElement } from "./xmlUtils";

// ============================================================================
// BORDER SERIALIZATION
// ============================================================================

/**
 * Serialize paragraph borders (w:pBdr)
 */
function serializeParagraphBorders(borders: ParagraphFormatting["borders"]): string {
  if (!borders) {
    return "";
  }

  const parts: string[] = [];

  if (borders.top) {
    const topXml = serializeBorder(borders.top, "top");
    if (topXml) {
      parts.push(topXml);
    }
  }

  if (borders.left) {
    const leftXml = serializeBorder(borders.left, "left");
    if (leftXml) {
      parts.push(leftXml);
    }
  }

  if (borders.bottom) {
    const bottomXml = serializeBorder(borders.bottom, "bottom");
    if (bottomXml) {
      parts.push(bottomXml);
    }
  }

  if (borders.right) {
    const rightXml = serializeBorder(borders.right, "right");
    if (rightXml) {
      parts.push(rightXml);
    }
  }

  if (borders.between) {
    const betweenXml = serializeBorder(borders.between, "between");
    if (betweenXml) {
      parts.push(betweenXml);
    }
  }

  if (borders.bar) {
    const barXml = serializeBorder(borders.bar, "bar");
    if (barXml) {
      parts.push(barXml);
    }
  }

  if (parts.length === 0) {
    return "";
  }

  return `<w:pBdr>${parts.join("")}</w:pBdr>`;
}

// ============================================================================
// SHADING SERIALIZATION
// ============================================================================

/**
 * Serialize shading properties (w:shd)
 */
function serializeShading(shading: ShadingProperties | undefined): string {
  if (!shading) {
    return "";
  }

  const attrs: string[] = [];

  // Pattern/val
  if (shading.pattern) {
    attrs.push(`w:val="${escapeXml(shading.pattern)}"`);
  } else {
    attrs.push('w:val="clear"');
  }

  // Color (pattern color)
  if (shading.color?.rgb && isValidHexColor(shading.color.rgb)) {
    attrs.push(`w:color="${escapeXml(shading.color.rgb)}"`);
  } else if (shading.color?.auto) {
    attrs.push('w:color="auto"');
  }

  // Fill (background color)
  if (shading.fill?.rgb && isValidHexColor(shading.fill.rgb)) {
    attrs.push(`w:fill="${escapeXml(shading.fill.rgb)}"`);
  } else if (shading.fill?.auto) {
    attrs.push('w:fill="auto"');
  }

  // Theme fill
  if (shading.fill?.themeColor) {
    attrs.push(`w:themeFill="${escapeXml(shading.fill.themeColor)}"`);
  }

  if (shading.fill?.themeTint) {
    attrs.push(`w:themeFillTint="${escapeXml(shading.fill.themeTint)}"`);
  }

  if (shading.fill?.themeShade) {
    attrs.push(`w:themeFillShade="${escapeXml(shading.fill.themeShade)}"`);
  }

  if (attrs.length === 0) {
    return "";
  }

  return `<w:shd ${attrs.join(" ")}/>`;
}

// ============================================================================
// TAB STOPS SERIALIZATION
// ============================================================================

/**
 * Serialize tab stops (w:tabs)
 */
function serializeTabStops(tabs: TabStop[] | undefined): string {
  if (!tabs || tabs.length === 0) {
    return "";
  }

  const tabElements = tabs.map((tab) => {
    const attrs: string[] = [`w:val="${tab.alignment}"`, `w:pos="${intAttr(tab.position)}"`];

    if (tab.leader && tab.leader !== "none") {
      attrs.push(`w:leader="${tab.leader}"`);
    }

    return `<w:tab ${attrs.join(" ")}/>`;
  });

  return `<w:tabs>${tabElements.join("")}</w:tabs>`;
}

// ============================================================================
// SPACING SERIALIZATION
// ============================================================================

/**
 * Serialize spacing properties (w:spacing)
 */
function serializeSpacing(formatting: ParagraphFormatting): string {
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

  if (attrs.length === 0) {
    return "";
  }

  return `<w:spacing ${attrs.join(" ")}/>`;
}

// ============================================================================
// INDENTATION SERIALIZATION
// ============================================================================

/**
 * Serialize indentation properties (w:ind)
 */
function serializeIndentation(formatting: ParagraphFormatting): string {
  const attrs: string[] = [];

  if (formatting.indentLeft !== undefined) {
    attrs.push(`w:left="${intAttr(formatting.indentLeft)}"`);
  }

  if (formatting.indentRight !== undefined) {
    attrs.push(`w:right="${intAttr(formatting.indentRight)}"`);
  }

  if (formatting.indentFirstLine !== undefined) {
    if (formatting.hangingIndent) {
      // Hanging indent is stored as positive value but uses w:hanging attribute
      attrs.push(`w:hanging="${intAttr(Math.abs(formatting.indentFirstLine))}"`);
    } else {
      attrs.push(`w:firstLine="${intAttr(formatting.indentFirstLine)}"`);
    }
  }

  if (attrs.length === 0) {
    return "";
  }

  return `<w:ind ${attrs.join(" ")}/>`;
}

// ============================================================================
// NUMBERING SERIALIZATION
// ============================================================================

/**
 * Serialize numbering properties (w:numPr)
 */
function serializeNumbering(numPr: ParagraphFormatting["numPr"]): string {
  if (!numPr) {
    return "";
  }

  const parts: string[] = [];

  if (numPr.ilvl !== undefined) {
    parts.push(`<w:ilvl w:val="${intAttr(numPr.ilvl)}"/>`);
  }

  if (numPr.numId !== undefined) {
    parts.push(`<w:numId w:val="${intAttr(numPr.numId)}"/>`);
  }

  if (parts.length === 0) {
    return "";
  }

  return `<w:numPr>${parts.join("")}</w:numPr>`;
}

// ============================================================================
// FRAME PROPERTIES SERIALIZATION
// ============================================================================

/**
 * Serialize frame properties (w:framePr)
 */
function serializeFrameProperties(frame: ParagraphFormatting["frame"]): string {
  if (!frame) {
    return "";
  }

  const attrs: string[] = [];

  if (frame.dropCap) {
    attrs.push(`w:dropCap="${frame.dropCap}"`);
  }

  if (frame.lines !== undefined) {
    attrs.push(`w:lines="${intAttr(frame.lines)}"`);
  }

  if (frame.width !== undefined) {
    attrs.push(`w:w="${intAttr(frame.width)}"`);
  }

  if (frame.height !== undefined) {
    attrs.push(`w:h="${intAttr(frame.height)}"`);
  }

  if (frame.hSpace !== undefined) {
    attrs.push(`w:hSpace="${intAttr(frame.hSpace)}"`);
  }

  if (frame.vSpace !== undefined) {
    attrs.push(`w:vSpace="${intAttr(frame.vSpace)}"`);
  }

  if (frame.hAnchor) {
    attrs.push(`w:hAnchor="${frame.hAnchor}"`);
  }

  if (frame.vAnchor) {
    attrs.push(`w:vAnchor="${frame.vAnchor}"`);
  }

  if (frame.x !== undefined) {
    attrs.push(`w:x="${frame.x}"`);
  }

  if (frame.y !== undefined) {
    attrs.push(`w:y="${frame.y}"`);
  }

  if (frame.xAlign) {
    attrs.push(`w:xAlign="${frame.xAlign}"`);
  }

  if (frame.yAlign) {
    attrs.push(`w:yAlign="${frame.yAlign}"`);
  }

  if (frame.wrap) {
    attrs.push(`w:wrap="${frame.wrap}"`);
  }

  if (attrs.length === 0) {
    return "";
  }

  return `<w:framePr ${attrs.join(" ")}/>`;
}

// ============================================================================
// PARAGRAPH PROPERTIES SERIALIZATION
// ============================================================================

/**
 * Serialize paragraph formatting properties to w:pPr XML
 */
function serializeParagraphMarkChange(mark: ParagraphMarkChange): string {
  const attrs = serializeTrackedChangeAttributes(mark.info);
  return `<w:${mark.kind} ${attrs}/>`;
}

type SerializeParagraphFormattingOptions = {
  propertyChanges?: ParagraphPropertyChange[] | undefined;
  paragraphMarkChange?: ParagraphMarkChange | undefined;
  propertySource?: ParagraphPropertySource | undefined;
  sectionProperties?: SectionProperties | undefined;
};

type ParagraphPropertySource = NonNullable<ReturnType<typeof getParagraphPropertySource>>;

const RESERVED_PARAGRAPH_PROPERTY_CHILDREN = new Set(["pPrChange", "sectPr"]);
const RESERVED_PARAGRAPH_CAPTURE_CHILDREN: ReadonlySet<string> = new Set([
  ...RESERVED_PARAGRAPH_PROPERTY_CHILDREN,
  ...PARAGRAPH_MARK_CHANGE_KINDS,
  "cellDel",
  "cellIns",
  "cellMerge",
  "numberingChange",
  "rPrChange",
  "tblGridChange",
  "tblPrChange",
  "tcPrChange",
  "trPrChange",
]);
const PARAGRAPH_PROPERTY_ROOT_NAME = new Set(["pPr"]);
const WORDPROCESSINGML_NAMESPACE = new Set([NAMESPACES.w]);
const PARAGRAPH_APPEND_PREFIXES = new Map([
  ["w", NAMESPACES.w],
  ["r", NAMESPACES.r],
  ["w15", NAMESPACES.w15],
]);
const PARAGRAPH_PROPERTY_CHILD_ORDER = new Map(
  [
    "pStyle",
    "keepNext",
    "keepLines",
    "pageBreakBefore",
    "framePr",
    "widowControl",
    "numPr",
    "suppressLineNumbers",
    "pBdr",
    "shd",
    "tabs",
    "suppressAutoHyphens",
    "kinsoku",
    "wordWrap",
    "overflowPunct",
    "topLinePunct",
    "autoSpaceDE",
    "autoSpaceDN",
    "bidi",
    "adjustRightInd",
    "snapToGrid",
    "spacing",
    "ind",
    "contextualSpacing",
    "mirrorIndents",
    "suppressOverlap",
    "jc",
    "textDirection",
    "textAlignment",
    "textboxTightWrap",
    "outlineLvl",
    "divId",
    "cnfStyle",
    "rPr",
  ].map((name, index) => [name, index]),
);
const PARAGRAPH_MARK_BASE_CHILDREN: ReadonlySet<string> = new Set([
  "rStyle",
  "rFonts",
  "b",
  "bCs",
  "i",
  "iCs",
  "caps",
  "smallCaps",
  "strike",
  "dstrike",
  "outline",
  "shadow",
  "emboss",
  "imprint",
  "snapToGrid",
  "vanish",
  "webHidden",
  "color",
  "spacing",
  "w",
  "kern",
  "position",
  "sz",
  "szCs",
  "noProof",
  "highlight",
  "u",
  "effect",
  "bdr",
  "shd",
  "fitText",
  "vertAlign",
  "rtl",
  "cs",
  "em",
  "lang",
  "eastAsianLayout",
  "specVanish",
  "oMath",
]);
const PARAGRAPH_NESTED_PROPERTY_CHILDREN: ReadonlyMap<string, ReadonlySet<string>> = new Map([
  ["numPr", new Set(["ilvl", "numId"])],
  ["pBdr", new Set(["top", "left", "bottom", "right", "between", "bar"])],
  ["tabs", new Set(["tab"])],
  ["rPr", PARAGRAPH_MARK_BASE_CHILDREN],
]);

const resolveNamespaceBinding = (
  scope: XmlNamespaceScope | undefined,
  prefix: string,
): string | undefined => {
  for (let current = scope; current; current = current.parent) {
    const value = current.bindings.get(prefix);
    if (value !== undefined) {
      return value;
    }
  }
  return undefined;
};

const hasDescendantNamed = (
  element: XmlElement,
  namespaceUri: string,
  localName: string,
): boolean => {
  const pending = [...getChildElements(element)];
  while (pending.length > 0) {
    const child = pending.pop();
    if (!child) {
      continue;
    }
    if (
      getLocalName(child.name) === localName &&
      toTransitionalNamespaceUri(child.namespaceUri ?? "") === namespaceUri
    ) {
      return true;
    }
    for (const descendant of getChildElements(child)) {
      pending.push(descendant);
    }
  }
  return false;
};

const hasInvalidParagraphPropertyRevision = (root: XmlElement): boolean => {
  const pending = [...getChildElements(root)];
  while (pending.length > 0) {
    const element = pending.pop();
    if (!element) {
      continue;
    }
    if (toTransitionalNamespaceUri(element.namespaceUri ?? "") === NAMESPACES.w) {
      const localName = getLocalName(element.name);
      if (RESERVED_PARAGRAPH_CAPTURE_CHILDREN.has(localName)) {
        return true;
      }
    }
    for (const child of getChildElements(element)) {
      pending.push(child);
    }
  }
  return false;
};

const hasInvalidParagraphPropertyShape = (root: XmlElement): boolean => {
  const pending = getChildElements(root).map((element) => ({ element, parent: root }));
  if (
    (root.elements ?? []).some(
      (child) => child.type === "text" && String(child.text ?? "").trim() !== "",
    )
  ) {
    return true;
  }
  while (pending.length > 0) {
    const entry = pending.pop();
    if (!entry) {
      continue;
    }
    const { element, parent } = entry;
    if (
      (element.elements ?? []).some(
        (child) => child.type === "text" && String(child.text ?? "").trim() !== "",
      )
    ) {
      return true;
    }
    const isWordprocessingElement =
      toTransitionalNamespaceUri(element.namespaceUri ?? "") === NAMESPACES.w;
    if (isWordprocessingElement && parent !== root) {
      if (toTransitionalNamespaceUri(parent.namespaceUri ?? "") !== NAMESPACES.w) {
        return true;
      }
      const allowedChildren = PARAGRAPH_NESTED_PROPERTY_CHILDREN.get(getLocalName(parent.name));
      if (!allowedChildren?.has(getLocalName(element.name))) {
        return true;
      }
    }
    for (const child of getChildElements(element)) {
      pending.push({ element: child, parent: element });
    }
  }
  return false;
};

const sourceShadowsDateUtcPrefix = (sourceXml: string): boolean => {
  const root = parseXml(sourceXml, OOXML_NAMESPACE_SCOPE).elements?.at(0);
  if (!root || root.type !== "element") {
    return true;
  }
  const pending = [root];
  while (pending.length > 0) {
    const element = pending.pop();
    if (!element) {
      continue;
    }
    const resolved = resolveNamespaceBinding(element.namespaceScope, "w16du");
    if (resolved !== undefined && resolved !== DATE_UTC_NAMESPACE_URI) {
      return true;
    }
    for (const child of getChildElements(element)) {
      pending.push(child);
    }
  }
  return false;
};

const hasInvalidParagraphMarkProperties = (root: XmlElement): boolean => {
  for (const child of getChildElements(root)) {
    if (toTransitionalNamespaceUri(child.namespaceUri ?? "") !== NAMESPACES.w) {
      continue;
    }
    const localName = getLocalName(child.name);
    if (!PARAGRAPH_MARK_BASE_CHILDREN.has(localName)) {
      return true;
    }
  }
  return false;
};

const replayableParagraphPropertySourceXml = (sourceXml: string): string | null => {
  return sanitizeCapturedXmlElement(sourceXml, {
    allowedLocalNames: PARAGRAPH_PROPERTY_ROOT_NAME,
    allowedNamespaceUris: WORDPROCESSINGML_NAMESPACE,
    inheritedNamespaceScope: OOXML_NAMESPACE_SCOPE,
    requiredNamespaceBindings: PARAGRAPH_APPEND_PREFIXES,
    validate: (root) => {
      let paragraphMarkProperties: XmlElement | null = null;
      const seenChildren = new Set<string>();
      let previousChildOrder = -1;
      for (const child of getChildElements(root)) {
        const localName = getLocalName(child.name);
        const isWordprocessingChild =
          toTransitionalNamespaceUri(child.namespaceUri ?? "") === NAMESPACES.w;
        if (isWordprocessingChild) {
          if (RESERVED_PARAGRAPH_PROPERTY_CHILDREN.has(localName)) {
            return false;
          }
          const childOrder = PARAGRAPH_PROPERTY_CHILD_ORDER.get(localName);
          if (
            childOrder === undefined ||
            seenChildren.has(localName) ||
            childOrder < previousChildOrder
          ) {
            return false;
          }
          seenChildren.add(localName);
          previousChildOrder = childOrder;
        }
        if (localName !== "rPr" || !isWordprocessingChild) {
          if (hasDescendantNamed(child, NAMESPACES.w, "rPr")) {
            return false;
          }
          continue;
        }
        if (
          paragraphMarkProperties !== null ||
          toTransitionalNamespaceUri(resolveNamespaceBinding(child.namespaceScope, "w") ?? "") !==
            NAMESPACES.w
        ) {
          return false;
        }
        paragraphMarkProperties = child;
      }
      return !(
        (paragraphMarkProperties && hasInvalidParagraphMarkProperties(paragraphMarkProperties)) ||
        hasInvalidParagraphPropertyShape(root) ||
        hasInvalidParagraphPropertyRevision(root)
      );
    },
  });
};

const verifiedParagraphPropertySource = (
  formatting: ParagraphFormatting | undefined,
  source: ParagraphPropertySource | undefined,
): string | null => {
  if (!source) {
    return null;
  }
  const replayableSource = replayableParagraphPropertySourceXml(source.xml);
  if (replayableSource === null) {
    return null;
  }
  return canonicalParagraphPropertySourceFingerprintJson(
    paragraphPropertySourceFingerprintFromFormatting(formatting),
  ) === source.fingerprintJson
    ? replayableSource
    : null;
};

const withTrailingParagraphPropertyChildren = (
  sourceXml: string,
  children: readonly string[],
): string => {
  const written = children.join("");
  if (written.length === 0) {
    return sourceXml;
  }
  const selfClosing = /^<((?:[A-Za-z_][\w.-]*:)?pPr)\b([^>]*)\/>$/u.exec(sourceXml);
  if (selfClosing) {
    const [, name, attrs] = selfClosing;
    return `<${name}${attrs}>${written}</${name}>`;
  }
  const close = sourceXml.lastIndexOf("</");
  return close === -1
    ? sourceXml
    : `${sourceXml.slice(0, close)}${written}${sourceXml.slice(close)}`;
};

const withParagraphMarkChange = (sourceXml: string, mark: ParagraphMarkChange): string => {
  const root = parseXml(sourceXml, OOXML_NAMESPACE_SCOPE).elements?.at(0);
  if (!root || root.type !== "element") {
    panic("A validated paragraph-property capture could not be parsed for composition");
  }
  const attributes = trackedChangeAttributeRecord(mark.info);
  const markElement = cloneElement(root, {
    name: `w:${mark.kind}`,
    attributes,
    elements: [],
  });
  const elements = [...(root.elements ?? [])];
  const paragraphMarkIndex = elements.findIndex(
    (child) =>
      child.type === "element" &&
      getLocalName(child.name) === "rPr" &&
      toTransitionalNamespaceUri(child.namespaceUri ?? "") === NAMESPACES.w,
  );
  if (paragraphMarkIndex === -1) {
    elements.push(
      cloneElement(root, {
        name: "w:rPr",
        attributes: {},
        elements: [markElement],
      }),
    );
  } else {
    const paragraphMarkProperties = elements[paragraphMarkIndex];
    if (!paragraphMarkProperties || paragraphMarkProperties.type !== "element") {
      panic("A validated paragraph-mark property node disappeared during composition");
    }
    elements[paragraphMarkIndex] = cloneElement(paragraphMarkProperties, {
      elements: [markElement, ...(paragraphMarkProperties.elements ?? [])],
    });
  }
  return captureVerbatimXml(cloneElement(root, { elements }));
};

const serializeParagraphFormattingWithOptions = (
  formatting: ParagraphFormatting | undefined,
  {
    propertyChanges,
    paragraphMarkChange,
    propertySource,
    sectionProperties,
  }: SerializeParagraphFormattingOptions = {},
): string => {
  // Suggested editor-only history is stripped before it reaches the document
  // model, so every entry here would serialize as a sibling w:pPrChange.
  const serializablePropertyChangeCount = propertyChanges?.length ?? 0;
  if (serializablePropertyChangeCount > 1) {
    panic("A paragraph cannot serialize more than one w:pPrChange", {
      count: serializablePropertyChangeCount,
    });
  }
  const paragraphMarkXml = paragraphMarkChange
    ? serializeParagraphMarkChange(paragraphMarkChange)
    : "";
  const sectionPropertiesXml = serializeSectionProperties(sectionProperties);
  const propertyChangesXml = (propertyChanges ?? []).map((change) =>
    serializeParagraphPropertyChange(change),
  );
  const composedChildrenUseDateUtc = [
    paragraphMarkXml,
    sectionPropertiesXml,
    ...propertyChangesXml,
  ].some((xml) => xml.includes(`${DATE_UTC_ATTRIBUTE}=`));
  const verifiedSource = verifiedParagraphPropertySource(formatting, propertySource);
  if (
    verifiedSource !== null &&
    (!composedChildrenUseDateUtc || !sourceShadowsDateUtcPrefix(verifiedSource))
  ) {
    const sourceWithMark = paragraphMarkChange
      ? withParagraphMarkChange(verifiedSource, paragraphMarkChange)
      : verifiedSource;
    return withTrailingParagraphPropertyChildren(sourceWithMark, [
      sectionPropertiesXml,
      ...propertyChangesXml,
    ]);
  }

  const parts: string[] = [];

  // Emit a boolean toggle: a bare element for true, `w:val="0"` for an explicit
  // false (which disables a value inherited from a style, so the override
  // survives round-trip), and nothing when absent ("inherit").
  const pushToggle = (name: string, value: boolean | undefined): void => {
    if (value === true) {
      parts.push(`<w:${name}/>`);
    } else if (value === false) {
      parts.push(`<w:${name} w:val="0"/>`);
    }
  };

  if (formatting) {
    // Style reference (must be first)
    if (formatting.styleId) {
      parts.push(`<w:pStyle w:val="${escapeXml(formatting.styleId)}"/>`);
    }

    // `CT_PPrBase` is a SEQUENCE, so every child below is written where that
    // sequence puts it, not where it reads best: a validating consumer that
    // meets one out of order reports the NEXT element as unexpected and
    // refuses the part. `contextualSpacing` and `snapToGrid` sit after the
    // indentation and the spacing, several elements past `numPr`, which is
    // what an early one made unexpected.
    pushToggle("keepNext", formatting.keepNext);
    pushToggle("keepLines", formatting.keepLines);
    pushToggle("pageBreakBefore", formatting.pageBreakBefore);

    // Frame properties
    const frameXml = serializeFrameProperties(formatting.frame);
    if (frameXml) {
      parts.push(frameXml);
    }

    // Widow control
    pushToggle("widowControl", formatting.widowControl);

    // Numbering. Skip numPr that still equals its style-sourced value (see
    // ParagraphFormatting.numPrFromStyle) — the parser materialized it from
    // the style and writing it back as direct formatting would flip Word's
    // level-indent precedence on the saved file. Guards the direct
    // serialize-a-parsed-Document path; the PM save path already drops it
    // in fromProseDoc.
    const styleSourcedNumPr =
      formatting.numPrFromStyle != null && numPrEqual(formatting.numPr, formatting.numPrFromStyle);
    const numPrXml = styleSourcedNumPr ? "" : serializeNumbering(formatting.numPr);
    if (numPrXml) {
      parts.push(numPrXml);
    }

    // Suppress line numbers precedes borders in CT_PPrBase.
    pushToggle("suppressLineNumbers", formatting.suppressLineNumbers);

    // Paragraph borders
    const bordersXml = serializeParagraphBorders(formatting.borders);
    if (bordersXml) {
      parts.push(bordersXml);
    }

    // Shading
    const shadingXml = serializeShading(formatting.shading);
    if (shadingXml) {
      parts.push(shadingXml);
    }

    // Tabs
    const tabsXml = serializeTabStops(formatting.tabs);
    if (tabsXml) {
      parts.push(tabsXml);
    }

    // Auto hyphens
    pushToggle("suppressAutoHyphens", formatting.suppressAutoHyphens);
    pushToggle("kinsoku", formatting.kinsoku);
    pushToggle("overflowPunct", formatting.overflowPunctuation);

    // Text direction (bidi)
    pushToggle("bidi", formatting.bidi);

    pushToggle("snapToGrid", formatting.snapToGrid);

    // Spacing
    const spacingXml = serializeSpacing(formatting);
    if (spacingXml) {
      parts.push(spacingXml);
    }

    // Indentation
    const indXml = serializeIndentation(formatting);
    if (indXml) {
      parts.push(indXml);
    }

    pushToggle("contextualSpacing", formatting.contextualSpacing);

    // Justification
    if (formatting.alignment) {
      parts.push(`<w:jc w:val="${formatting.alignment}"/>`);
    }

    // Outline level
    if (formatting.outlineLevel !== undefined) {
      parts.push(`<w:outlineLvl w:val="${formatting.outlineLevel}"/>`);
    }

    // Run properties (default run formatting for paragraph)
    // Round-trip `<w:specVanish/>` (run-in heading marker, ECMA-376
    // §17.3.1.32) by injecting it into the paragraph mark's rPr.
    // The parser populates `formatting.runInWithNext` from this
    // element; the layout engine consumes it via toFlowBlocks'
    // run-in merge. Without serializing it back, saving a doc
    // through Folio loses the soft paragraph break and the heading
    // becomes a normal separate paragraph in Word.
    //
    // EG_ParaRPrTrackChanges (ECMA-376 §17.13.5 / wml.xsd:1837) puts
    // <w:ins>/<w:del> FIRST inside the paragraph mark's rPr; strict
    // readers reject other orderings.
    if (paragraphMarkChange || formatting.runProperties || formatting.runInWithNext) {
      const pPrMarkXml = paragraphMarkChange ? paragraphMarkXml : "";
      const innerRPr = formatting.runProperties
        ? extractRPrInner(serializeTextFormatting(formatting.runProperties))
        : "";
      const specVanishXml = formatting.runInWithNext ? "<w:specVanish/>" : "";
      const fullInner = `${pPrMarkXml}${innerRPr}${specVanishXml}`;
      if (fullInner.length > 0) {
        parts.push(`<w:rPr>${fullInner}</w:rPr>`);
      }
    }
  } else if (paragraphMarkChange) {
    parts.push(`<w:rPr>${paragraphMarkXml}</w:rPr>`);
  }

  // `CT_PPr` closes with `rPr`, `sectPr`, `pPrChange` in that order: a section
  // break sits between the mark's run properties and the recorded change, so
  // it is placed here rather than appended after the properties are built.
  parts.push(sectionPropertiesXml);

  if (propertyChangesXml.length > 0) {
    parts.push(...propertyChangesXml);
  }

  const inner = parts.join("");
  if (inner.length === 0) {
    return "";
  }

  return `<w:pPr>${inner}</w:pPr>`;
};

export function serializeParagraphFormatting(
  formatting: ParagraphFormatting | undefined,
  propertyChanges?: ParagraphPropertyChange[],
  paragraphMarkChange?: ParagraphMarkChange,
): string {
  return serializeParagraphFormattingWithOptions(formatting, {
    propertyChanges,
    paragraphMarkChange,
  });
}

function extractPPrInner(pPrXml: string): string {
  if (!pPrXml.startsWith("<w:pPr>") || !pPrXml.endsWith("</w:pPr>")) {
    return "";
  }
  return pPrXml.slice("<w:pPr>".length, -"</w:pPr>".length);
}

/**
 * Strip the outer `<w:rPr>...</w:rPr>` wrapper so callers can splice
 * additional rPr children (e.g. `<w:specVanish/>`) and re-emit a
 * single rPr element.
 */
function extractRPrInner(rPrXml: string): string {
  if (!rPrXml.startsWith("<w:rPr>") || !rPrXml.endsWith("</w:rPr>")) {
    return "";
  }
  return rPrXml.slice("<w:rPr>".length, -"</w:rPr>".length);
}

function serializeParagraphPropertyChange(change: ParagraphPropertyChange): string {
  const previousPPrXml = serializeParagraphFormatting(change.previousFormatting) || "<w:pPr/>";
  const previousPPrInner = extractPPrInner(previousPPrXml);
  const normalizedPreviousPPr =
    previousPPrInner.length > 0 ? `<w:pPr>${previousPPrInner}</w:pPr>` : "<w:pPr/>";
  return `<w:pPrChange ${serializeTrackedChangeAttributes(change.info)}>${normalizedPreviousPPr}</w:pPrChange>`;
}

// ============================================================================
// CONTENT SERIALIZATION
// ============================================================================

/** The attribute list of a `w:hyperlink`, without its children. */
function hyperlinkAttributes(hyperlink: Hyperlink): string {
  const attrs: string[] = [];

  if (hyperlink.rId) {
    attrs.push(`r:id="${escapeXml(hyperlink.rId)}"`);
  }

  if (hyperlink.anchor) {
    attrs.push(`w:anchor="${escapeXml(hyperlink.anchor)}"`);
  }

  if (hyperlink.tooltip) {
    attrs.push(`w:tooltip="${escapeXml(hyperlink.tooltip)}"`);
  }

  if (hyperlink.target) {
    attrs.push(`w:tgtFrame="${escapeXml(hyperlink.target)}"`);
  }

  // Round-trip an explicit `w:history` either way. The parser only sets
  // `history` from a present `w:history="1"`/`"0"`, so emitting nothing for
  // `true` used to drop the attribute on save.
  if (hyperlink.history === true) {
    attrs.push('w:history="1"');
  } else if (hyperlink.history === false) {
    attrs.push('w:history="0"');
  }

  if (hyperlink.docLocation) {
    attrs.push(`w:docLocation="${escapeXml(hyperlink.docLocation)}"`);
  }

  return attrs.length > 0 ? ` ${attrs.join(" ")}` : "";
}

/** One `w:hyperlink` child, with the caller deciding how a run is written. */
function serializeHyperlinkChild(
  child: Hyperlink["children"][number],
  serializeChildRun: (run: Run) => string,
): string {
  if (child.type === "run") {
    return serializeChildRun(child);
  }
  return child.type === "bookmarkStart"
    ? serializeBookmarkStart(child)
    : serializeBookmarkEnd(child);
}

/**
 * Serialize a hyperlink (w:hyperlink)
 */
function serializeHyperlink(hyperlink: Hyperlink): string {
  const childrenXml = hyperlink.children
    .map((child) => serializeHyperlinkChild(child, serializeRun))
    .join("");
  return `<w:hyperlink${hyperlinkAttributes(hyperlink)}>${childrenXml}</w:hyperlink>`;
}

/**
 * Serialize bookmark start (w:bookmarkStart)
 */
function serializeBookmarkStart(bookmark: BookmarkStart): string {
  const attrs: string[] = [`w:id="${bookmark.id}"`, `w:name="${escapeXml(bookmark.name)}"`];

  if (bookmark.colFirst !== undefined) {
    attrs.push(`w:colFirst="${bookmark.colFirst}"`);
  }

  if (bookmark.colLast !== undefined) {
    attrs.push(`w:colLast="${bookmark.colLast}"`);
  }

  return `<w:bookmarkStart ${attrs.join(" ")}/>`;
}

/**
 * Serialize bookmark end (w:bookmarkEnd)
 */
function serializeBookmarkEnd(bookmark: BookmarkEnd): string {
  return `<w:bookmarkEnd w:id="${bookmark.id}"/>`;
}

/** Serialize a simple field without changing its authored OOXML field form. */
function serializeSimpleField(field: SimpleField): string {
  const attrs = [`w:instr="${escapeXml(field.instruction)}"`];
  if (field.fldLock) {
    attrs.push('w:fldLock="true"');
  }
  if (field.dirty) {
    attrs.push('w:dirty="true"');
  }

  const contentXml = field.content
    .map((item) => (item.type === "run" ? serializeRun(item) : serializeHyperlink(item)))
    .join("");

  return `<w:fldSimple ${attrs.join(" ")}>${contentXml}</w:fldSimple>`;
}

/**
 * Serialize a complex field
 * Complex fields are represented by multiple runs with fldChar elements,
 * so we convert them back to that structure
 */
function serializeComplexField(field: ComplexField): string {
  const parts: string[] = [];

  // Formatting for the structural runs (begin/separate/end). Prefer the field's
  // captured run formatting: the parser re-captures ComplexField.formatting from
  // the first non-empty field-run `w:rPr`, so re-presenting it on the begin run
  // keeps that value stable across a save→parse round-trip. Fall back to the
  // first result run's formatting when the field has no captured formatting.
  // The collapsed PAGE field (no result run) still recovers its `w:rPr`
  // (size/color) from field.formatting (eigenpal/docx-editor#909).
  const structuralFormatting = field.formatting ?? field.fieldResult[0]?.formatting;
  const rPrXml = structuralFormatting ? serializeTextFormatting(structuralFormatting) : "";

  // Begin field character. `dirty` is emitted only when the model asks for it:
  // it makes consumers recompute the field on open (and may discard result
  // run formatting), which is what a generated TOC wants and nothing else.
  const beginAttrs: string[] = ['w:fldCharType="begin"'];
  if (field.fldLock) {
    beginAttrs.push('w:fldLock="true"');
  }
  if (field.dirty) {
    beginAttrs.push('w:dirty="true"');
  }
  parts.push(`<w:r>${rPrXml}<w:fldChar ${beginAttrs.join(" ")}/></w:r>`);

  // Field code (instrText)
  if (field.fieldCode.length > 0) {
    parts.push(...field.fieldCode.map((run) => serializeRun(run)));
  } else if (field.instruction.length > 0) {
    // Fallback: create instrText from instruction
    const needsPreserve =
      field.instruction.startsWith(" ") ||
      field.instruction.endsWith(" ") ||
      field.instruction.includes("  ");
    const spaceAttr = needsPreserve ? ' xml:space="preserve"' : "";
    parts.push(
      `<w:r>${rPrXml}<w:instrText${spaceAttr}>${escapeXml(field.instruction)}</w:instrText></w:r>`,
    );
  }

  // Separate field character
  parts.push(`<w:r>${rPrXml}<w:fldChar w:fldCharType="separate"/></w:r>`);

  // Field result
  parts.push(...field.fieldResult.map((run) => serializeRun(run)));

  // End field character
  parts.push(`<w:r>${rPrXml}<w:fldChar w:fldCharType="end"/></w:r>`);

  return parts.join("");
}

/**
 * Synthesize a `<w:sdtPr>` from the modeled {@link SdtProperties}.
 *
 * Only reached for an inline SDT that carries no captured `rawPropertiesXml`
 * (constructed programmatically rather than parsed from a DOCX). Mirrors the
 * block-SDT fallback: emit `w:id` first so the parsed numeric id survives,
 * then the shared identity fields, then the type-defining marker.
 */
function synthesizeInlineSdtPr(props: SdtProperties): string {
  const prParts: string[] = [];

  if (typeof props.id === "number") {
    prParts.push(`<w:id w:val="${props.id}"/>`);
  }
  if (props.alias) {
    prParts.push(`<w:alias w:val="${escapeXml(props.alias)}"/>`);
  }
  if (props.tag) {
    prParts.push(`<w:tag w:val="${escapeXml(props.tag)}"/>`);
  }
  if (props.lock && props.lock !== "unlocked") {
    prParts.push(`<w:lock w:val="${props.lock}"/>`);
  }
  if (props.placeholder) {
    // OOXML shape: `<w:placeholder><w:docPart w:val="..."/></w:placeholder>`.
    // The placeholder identifier lives in `w:val` on the nested `w:docPart`,
    // mirroring the parse in `paragraphParser.ts`.
    prParts.push(
      `<w:placeholder><w:docPart w:val="${escapeXml(props.placeholder)}"/></w:placeholder>`,
    );
  }
  if (props.showingPlaceholder) {
    prParts.push("<w:showingPlcHdr/>");
  }

  // Type-specific properties
  switch (props.sdtType) {
    case "plainText":
      prParts.push("<w:text/>");
      break;
    case "date": {
      // `w:date@w:fullDate` is the ISO-8601 bound value; `w:dateFormat` is
      // the display format. Older code (before the shared parser
      // split these) wrote the format into `w:fullDate`, which corrupted
      // round-trip — keep them on separate model fields and emit each
      // into its right element.
      const fullDateAttr = props.dateValueISO
        ? ` w:fullDate="${escapeXml(props.dateValueISO)}"`
        : "";
      const formatChild = props.dateFormat
        ? `<w:dateFormat w:val="${escapeXml(props.dateFormat)}"/>`
        : "";
      if (fullDateAttr || formatChild) {
        prParts.push(`<w:date${fullDateAttr}>${formatChild}</w:date>`);
      } else {
        prParts.push("<w:date/>");
      }
      break;
    }
    case "dropdown": {
      const items = (props.listItems ?? [])
        .map(
          (i) =>
            `<w:listItem w:displayText="${escapeXml(i.displayText)}" w:value="${escapeXml(i.value)}"/>`,
        )
        .join("");
      prParts.push(`<w:dropDownList>${items}</w:dropDownList>`);
      break;
    }
    case "comboBox": {
      const items = (props.listItems ?? [])
        .map(
          (i) =>
            `<w:listItem w:displayText="${escapeXml(i.displayText)}" w:value="${escapeXml(i.value)}"/>`,
        )
        .join("");
      prParts.push(`<w:comboBox>${items}</w:comboBox>`);
      break;
    }
    case "checkbox":
      prParts.push(
        `<w14:checkbox><w14:checked w14:val="${props.checked ? "1" : "0"}"/></w14:checkbox>`,
      );
      break;
    case "picture":
      prParts.push("<w:picture/>");
      break;
    case "richText":
    case "buildingBlockGallery":
    case "group":
    case "unknown":
      // These SDT variants carry no type-specific properties in OOXML;
      // the surrounding sdtPr fields (alias/tag/lock/...) carry all
      // round-trippable state for them.
      break;
  }

  return `<w:sdtPr>${prParts.join("")}</w:sdtPr>`;
}

/**
 * Serialize an inline SDT (w:sdt).
 *
 * Replays the captured `<w:sdtPr>` / `<w:sdtEndPr>` verbatim so unmodeled
 * OOXML features (`w:id`, `w:dataBinding`, `w15:*`, custom XML mappings)
 * survive the round-trip, mirroring the block-SDT serializer. Without this
 * the properties block was re-synthesized from the modeled projection alone,
 * silently dropping every unmodeled feature (and `w:sdtEndPr`) on save.
 */
function serializeInlineSdt(sdt: InlineSdt): string {
  const props = sdt.properties;

  const contentXml = sdt.content
    .map((item): string => {
      switch (item.type) {
        case "run":
          return serializeRun(item);
        case "hyperlink":
          return serializeHyperlink(item);
        case "simpleField":
          return serializeSimpleField(item);
        case "complexField":
          return serializeComplexField(item);
        case "inlineSdt":
          return serializeInlineSdt(item);
        case "insertion":
          return serializeTrackedChange("ins", item);
        case "deletion":
          return serializeTrackedChange("del", item);
        case "moveFrom":
          return serializeTrackedChange("moveFrom", item);
        case "moveTo":
          return serializeTrackedChange("moveTo", item);
        case "mathEquation":
          // Round-trip the raw OMML XML directly
          return item.ommlXml || "";
        default: {
          // Exhaustiveness check: if a new type is added to
          // InlineSdt['content'] (see docx-core/src/model/content.ts)
          // without a matching case here, TypeScript errors out instead
          // of silently dropping content on save. Keep this in sync with
          // the filter in createInlineSdtFromNode (fromProseDoc.ts).
          const _exhaustive: never = item;
          return _exhaustive;
        }
      }
    })
    .join("");

  // Reconcile any modeled interactive edit (checkbox toggle, date pick,
  // dropdown selection) into the raw properties before replay so it is not
  // discarded, exactly as the block-SDT serializer does. Unmodeled markers
  // inside the raw string are left untouched.
  // Replay the captured snapshot only when it is structurally a single
  // `<w:sdtPr>`/`<w:sdtEndPr>` element — a malformed or attacker-supplied
  // string (e.g. one that closes `<w:sdt>` early or injects sibling markup)
  // falls back to a synthesized properties block instead of being spliced
  // into the document verbatim.
  const baseSdtPr =
    props.rawPropertiesXml && isSingleWellFormedElement(props.rawPropertiesXml, "sdtPr")
      ? props.rawPropertiesXml
      : synthesizeInlineSdtPr(props);
  const dateFullDate =
    props.sdtType === "date" && props.dateValueISO ? props.dateValueISO : undefined;
  const dropdownLastValue =
    (props.sdtType === "dropdown" || props.sdtType === "comboBox") &&
    typeof props.dropdownLastValue === "string"
      ? props.dropdownLastValue
      : undefined;
  const sdtPrXml = reconcileRawSdtPr(baseSdtPr, props, {
    ...(dateFullDate !== undefined ? { dateFullDate } : {}),
    ...(dropdownLastValue !== undefined ? { dropdownLastValue } : {}),
  });
  const sdtEndPrXml =
    props.rawEndPropertiesXml && isSingleWellFormedElement(props.rawEndPropertiesXml, "sdtEndPr")
      ? props.rawEndPropertiesXml
      : "";

  return `<w:sdt>${sdtPrXml}${sdtEndPrXml}<w:sdtContent>${contentXml}</w:sdtContent></w:sdt>`;
}

function serializeMoveRangeStart(
  tag: "moveFromRangeStart" | "moveToRangeStart",
  marker: MoveFromRangeStart | MoveToRangeStart,
): string {
  const attrs = [`w:id="${marker.id}"`, `w:name="${escapeXml(marker.name)}"`];
  return `<w:${tag} ${attrs.join(" ")}/>`;
}

/**
 * Serialize a tracked change wrapper (ins/del/moveFrom/moveTo)
 */
function rewriteRunTextAsDeleted(xml: string): string {
  return xml
    .replace(/<w:t\b/gu, "<w:delText")
    .replace(/<\/w:t>/gu, "</w:delText>")
    .replace(/<w:instrText\b/gu, "<w:delInstrText")
    .replace(/<\/w:instrText>/gu, "</w:delInstrText>");
}

function trackedChangeTag(
  change: Insertion | Deletion | MoveFrom | MoveTo,
): "ins" | "del" | "moveFrom" | "moveTo" {
  switch (change.type) {
    case "insertion":
      return "ins";
    case "deletion":
      return "del";
    case "moveFrom":
      return "moveFrom";
    case "moveTo":
      return "moveTo";
  }
}

function serializeTrackedChange(
  tag: "ins" | "del" | "moveFrom" | "moveTo",
  change: Insertion | Deletion | MoveFrom | MoveTo,
): string {
  const attrs = serializeTrackedChangeAttributes(change.info);

  const serializeDeletedRun = (run: Run): string => {
    const xml = serializeRun(run);
    const hasDrawingContent = run.content.some((c) => c.type === "drawing" || c.type === "shape");
    if (!hasDrawingContent) {
      return rewriteRunTextAsDeleted(xml);
    }

    const hasTextualContent = run.content.some((c) => c.type !== "drawing" && c.type !== "shape");
    if (!hasTextualContent) {
      return xml;
    }

    return run.content
      .map((content) => {
        const contentXml = serializeRun({ ...run, content: [content] });
        if (content.type === "drawing" || content.type === "shape") {
          return contentXml;
        }
        return rewriteRunTextAsDeleted(contentXml);
      })
      .join("");
  };

  const serializeContentRun = (run: Run): string =>
    // A deleted drawing/shape run keeps its content verbatim: a picture
    // has no `<w:t>`, and a shape's nested textbox text
    // (`<w:txbxContent><w:t>`) must NOT be rewritten to `<w:delText>` —
    // that markup belongs only to a run's own deleted text, not to a
    // nested textbox document. eigenpal #641.
    tag === "del" || tag === "moveFrom" ? serializeDeletedRun(run) : serializeRun(run);

  // A hyperlink is not written inside the wrapper at all, so it is not one of
  // the items this writes: the loop below opens the wrapper inside the link
  // instead.
  type WrappedItem = Exclude<(typeof change.content)[number], Hyperlink>;

  const serializeWrappedItem = (item: WrappedItem): string => {
    if (item.type === "run") {
      return serializeContentRun(item);
    }
    if (item.type === "simpleField" || item.type === "complexField") {
      const xml =
        item.type === "simpleField" ? serializeSimpleField(item) : serializeComplexField(item);
      return tag === "del" || tag === "moveFrom" ? rewriteRunTextAsDeleted(xml) : xml;
    }
    if (
      item.type === "insertion" ||
      item.type === "deletion" ||
      item.type === "moveFrom" ||
      item.type === "moveTo"
    ) {
      return serializeTrackedChange(trackedChangeTag(item), item);
    }
    return item.type === "bookmarkStart"
      ? serializeBookmarkStart(item)
      : serializeBookmarkEnd(item);
  };

  const open = `<w:${tag} ${attrs}>`;
  const close = `</w:${tag}>`;
  const wrap = (inner: string): string => (inner.length === 0 ? "" : `${open}${inner}${close}`);

  // An empty wrapper is a marker in its own right (a paragraph mark's
  // revision, a move end), so it survives the segmentation below.
  if (change.content.length === 0) {
    return `${open}${close}`;
  }

  // `w:hyperlink` may not appear inside a revision wrapper; the nesting runs
  // the other way, with the wrapper opened again around the linked runs. So a
  // revision spanning a hyperlink is emitted as several wrappers — text
  // before, the hyperlink carrying its own, text after — which the package's
  // revision-id pass then gives distinct `w:id`s.
  const segments: string[] = [];
  const pending: string[] = [];
  const flushPending = (): void => {
    if (pending.length > 0) {
      segments.push(wrap(pending.join("")));
      pending.length = 0;
    }
  };
  for (const item of change.content) {
    if (item.type === "hyperlink") {
      flushPending();
      const childrenXml = item.children
        .map((child) => serializeHyperlinkChild(child, serializeContentRun))
        .join("");
      // Always the full wrapper, never `wrap`: a linked run range that is
      // empty still has to say it was inserted or deleted, or reopening the
      // package finds a plain hyperlink.
      segments.push(
        `<w:hyperlink${hyperlinkAttributes(item)}>${open}${childrenXml}${close}</w:hyperlink>`,
      );
      continue;
    }
    pending.push(serializeWrappedItem(item));
  }
  flushPending();

  return segments.join("");
}

/** Emit the `<w:commentReference>` run Word places after a comment range end. */
function serializeCommentReferenceRun(id: number): string {
  return `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${id}"/></w:r>`;
}

/**
 * Serialize a single paragraph content item.
 *
 * `explicitCommentReferenceIds` holds the comment ids that already have their
 * own `commentReference` node in this paragraph (the parsed-from-Word shape).
 * For those, the `commentRangeEnd` marker must NOT also synthesize a reference
 * run, or a save→parse round-trip doubles the `<w:commentReference>`. The
 * editor (fromProseDoc) path emits range markers with no reference node, so the
 * synthetic run is still written when the id is absent from the set.
 */
function serializeParagraphContent(
  content: ParagraphContent,
  explicitCommentReferenceIds: ReadonlySet<number>,
): string {
  switch (content.type) {
    case "run":
      return serializeRun(content);
    case "hyperlink":
      return serializeHyperlink(content);
    case "bookmarkStart":
      return serializeBookmarkStart(content);
    case "bookmarkEnd":
      return serializeBookmarkEnd(content);
    case "simpleField":
      return serializeSimpleField(content);
    case "complexField":
      return serializeComplexField(content);
    case "inlineSdt":
      return serializeInlineSdt(content);
    case "commentRangeStart":
      return `<w:commentRangeStart w:id="${content.id}"/>`;
    case "commentRangeEnd":
      return explicitCommentReferenceIds.has(content.id)
        ? `<w:commentRangeEnd w:id="${content.id}"/>`
        : `<w:commentRangeEnd w:id="${content.id}"/>${serializeCommentReferenceRun(content.id)}`;
    case "commentReference":
      return serializeCommentReferenceRun(content.id);
    case "insertion":
      return serializeTrackedChange("ins", content);
    case "deletion":
      return serializeTrackedChange("del", content);
    case "moveFrom":
      return serializeTrackedChange("moveFrom", content);
    case "moveTo":
      return serializeTrackedChange("moveTo", content);
    case "moveFromRangeStart":
      return serializeMoveRangeStart("moveFromRangeStart", content as MoveFromRangeStart);
    case "moveFromRangeEnd":
      return `<w:moveFromRangeEnd w:id="${content.id}"/>`;
    case "moveToRangeStart":
      return serializeMoveRangeStart("moveToRangeStart", content as MoveToRangeStart);
    case "moveToRangeEnd":
      return `<w:moveToRangeEnd w:id="${content.id}"/>`;
    case "mathEquation":
      // Round-trip the raw OMML XML directly
      return content.ommlXml || "";
    default:
      return "";
  }
}

// ============================================================================
// MAIN SERIALIZATION
// ============================================================================

/**
 * Serialize a paragraph to OOXML XML (w:p)
 *
 * @param paragraph - The paragraph to serialize
 * @returns XML string for the paragraph
 */
export function serializeParagraph(paragraph: Paragraph): string {
  const parts: string[] = [];

  // Paragraph ID attributes
  const attrs: string[] = [];
  if (paragraph.paraId) {
    attrs.push(`w14:paraId="${escapeXml(paragraph.paraId)}"`);
  }
  if (paragraph.textId) {
    attrs.push(`w14:textId="${escapeXml(paragraph.textId)}"`);
  }
  const attrsStr = attrs.length > 0 ? ` ${attrs.join(" ")}` : "";

  // Add paragraph properties if present
  parts.push(
    serializeParagraphFormattingWithOptions(paragraph.formatting, {
      propertyChanges: paragraph.propertyChanges,
      paragraphMarkChange: paragraph.pPrMark,
      propertySource: getParagraphPropertySource(paragraph),
      sectionProperties: paragraph.sectionProperties,
    }),
  );

  // Comment ids whose reference run is modeled explicitly (parsed-from-Word),
  // so the matching commentRangeEnd does not double-emit it (see
  // serializeParagraphContent).
  const explicitCommentReferenceIds = new Set<number>();
  for (const content of paragraph.content) {
    if (content.type === "commentReference") {
      explicitCommentReferenceIds.add(content.id);
    }
  }

  // Add paragraph content
  let pendingRenderedPageBreak = paragraph.renderedPageBreakBefore === true;
  for (const content of paragraph.content) {
    let contentXml = serializeParagraphContent(content, explicitCommentReferenceIds);
    if (contentXml) {
      if (pendingRenderedPageBreak) {
        const next = injectRenderedPageBreakIntoFirstRun(contentXml);
        if (next) {
          contentXml = next;
          pendingRenderedPageBreak = false;
        }
      }
      parts.push(contentXml);
    }
  }

  return `<w:p${attrsStr}>${parts.join("")}</w:p>`;
}

function injectRenderedPageBreakIntoFirstRun(xml: string): string | null {
  const runOpeningTag = /<w:r(?=[\s>/])[^>]*>/u;
  const openingTag = runOpeningTag.exec(xml);
  if (!openingTag) {
    return null;
  }
  const runEnd = xml.indexOf("</w:r>", openingTag.index + openingTag[0].length);
  if (
    runEnd !== -1 &&
    xml
      .slice(openingTag.index + openingTag[0].length, runEnd)
      .includes("<w:lastRenderedPageBreak/>")
  ) {
    return xml;
  }
  const contentStart = openingTag.index + openingTag[0].length;
  let insertionOffset = contentStart;
  if (xml.startsWith("<w:rPr", contentStart)) {
    const propertyTag = /<\/?w:rPr(?=[\s>/])[^>]*>/gu;
    propertyTag.lastIndex = contentStart;
    let depth = 0;
    for (const match of xml.matchAll(propertyTag)) {
      if (match.index !== contentStart && depth === 0) {
        break;
      }
      if (match[0].startsWith("</")) {
        depth--;
      } else if (!match[0].endsWith("/>")) {
        depth++;
      }
      if (depth === 0) {
        insertionOffset = match.index + match[0].length;
        break;
      }
    }
  }
  return `${xml.slice(0, insertionOffset)}<w:lastRenderedPageBreak/>${xml.slice(insertionOffset)}`;
}

/**
 * Serialize multiple paragraphs to OOXML XML
 *
 * @param paragraphs - The paragraphs to serialize
 * @returns XML string for all paragraphs
 */
export function serializeParagraphs(paragraphs: Paragraph[]): string {
  return paragraphs.map(serializeParagraph).join("");
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/**
 * Check if a paragraph has any content
 */
export function hasParagraphContent(paragraph: Paragraph): boolean {
  return paragraph.content.length > 0;
}

/**
 * Check if a paragraph has formatting
 */
export function hasParagraphFormatting(paragraph: Paragraph): boolean {
  return paragraph.formatting !== undefined && Object.keys(paragraph.formatting).length > 0;
}

/**
 * Get plain text from a paragraph (for comparison/debugging)
 */
export function getParagraphPlainText(paragraph: Paragraph): string {
  const texts: string[] = [];

  for (const content of paragraph.content) {
    if (content.type === "run") {
      for (const item of content.content) {
        if (item.type === "text") {
          texts.push(item.text);
        } else if (item.type === "tab") {
          texts.push("\t");
        } else if (item.type === "break") {
          texts.push("\n");
        }
      }
    } else if (content.type === "hyperlink") {
      for (const child of content.children) {
        if (child.type === "run") {
          for (const item of child.content) {
            if (item.type === "text") {
              texts.push(item.text);
            }
          }
        }
      }
    } else if (
      content.type === "simpleField" ||
      content.type === "inlineSdt" ||
      content.type === "insertion" ||
      content.type === "deletion" ||
      content.type === "moveFrom" ||
      content.type === "moveTo"
    ) {
      for (const item of content.content) {
        if (item.type === "run") {
          for (const subItem of item.content) {
            if (subItem.type === "text") {
              texts.push(subItem.text);
            }
          }
        }
      }
    } else if (content.type === "complexField") {
      for (const run of content.fieldResult) {
        for (const item of run.content) {
          if (item.type === "text") {
            texts.push(item.text);
          }
        }
      }
    }
  }

  return texts.join("");
}

/**
 * Create an empty paragraph
 */
export function createEmptyParagraph(formatting?: ParagraphFormatting): Paragraph {
  return {
    type: "paragraph",
    ...(formatting !== undefined ? { formatting } : {}),
    content: [],
  };
}

/**
 * Create a paragraph with a single text run
 */
export function createTextParagraph(
  text: string,
  paragraphFormatting?: ParagraphFormatting,
  textFormatting?: TextFormatting,
): Paragraph {
  return {
    type: "paragraph",
    ...(paragraphFormatting !== undefined ? { formatting: paragraphFormatting } : {}),
    content: [
      {
        type: "run",
        ...(textFormatting !== undefined ? { formatting: textFormatting } : {}),
        content: [{ type: "text", text }],
      },
    ],
  };
}

/**
 * Check if paragraph is a list item
 */
export function isListParagraph(paragraph: Paragraph): boolean {
  return paragraph.formatting?.numPr !== undefined;
}

/**
 * Get list level of a paragraph (0-8, or -1 if not a list)
 */
export function getListLevel(paragraph: Paragraph): number {
  return paragraph.formatting?.numPr?.ilvl ?? -1;
}
