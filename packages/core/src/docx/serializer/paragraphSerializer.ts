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
  TextFormatting,
} from "../../types/document";
import { PARAGRAPH_MARK_CHANGE_KINDS } from "@stll/docx-core/model";
import { panic } from "better-result";
import {
  modelParagraphFormattingEmission,
  type ModeledParagraphFormattingEmission,
} from "../../internal/paragraphFormattingSerialization";
import {
  getParagraphPropertySource,
  paragraphPropertySourceMatchesEmission,
} from "../paragraphPropertySource";
import { fieldStateAttributes } from "../fieldState";
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
import {
  bookmarkRangeAttributes,
  markupRangeAttributes,
  moveBookmarkAttributes,
} from "./markupRangeAttributes";
// oxlint-disable-next-line import/no-cycle -- OOXML model is mutually recursive: paragraphs hold runs, shape-textbox runs hold paragraphs
import { serializeRun } from "./runSerializer";
import { serializeSectionProperties } from "./sectionPropertiesSerializer";
import { serializeTextFormatting } from "./textFormattingSerializer";
import {
  serializeTrackedChangeAttributes,
  trackedChangeAttributeRecord,
} from "./trackedChangeAttributes";
import { isSingleWellFormedElement } from "./xmlUtils";
import { escapeXmlAttribute, escapeXmlText } from "@stll/docx-core";

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
/**
 * Revision records a captured `w:pPr` may not carry into a replay.
 *
 * `w:numberingChange` is not among them: the model holds it
 * (`ParagraphFormatting.numberingChangeXml`) and the serializer writes it
 * back, so both paths keep it. Refusing the capture used to force a rebuild
 * that could not write it, which turned a replay gate into a lost revision.
 */
const RESERVED_PARAGRAPH_CAPTURE_CHILDREN: ReadonlySet<string> = new Set([
  ...RESERVED_PARAGRAPH_PROPERTY_CHILDREN,
  ...PARAGRAPH_MARK_CHANGE_KINDS,
  "cellDel",
  "cellIns",
  "cellMerge",
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
  ["numPr", new Set(["ilvl", "numId", "numberingChange"])],
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
  formatting: ModeledParagraphFormattingEmission,
  source: ParagraphPropertySource | undefined,
): string | null => {
  if (!source) {
    return null;
  }
  const replayableSource = replayableParagraphPropertySourceXml(source.xml);
  if (replayableSource === null) {
    return null;
  }
  return paragraphPropertySourceMatchesEmission(source, formatting) ? replayableSource : null;
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
  const modeledFormatting = modelParagraphFormattingEmission(formatting);
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
  const verifiedSource = verifiedParagraphPropertySource(modeledFormatting, propertySource);
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
  if (modeledFormatting.propertiesXml) {
    parts.push(modeledFormatting.propertiesXml);
  }
  // EG_ParaRPrTrackChanges puts revision markup first inside the paragraph
  // mark's rPr; modeled run properties and specVanish follow it.
  const paragraphMarkPropertiesInnerXml = `${paragraphMarkXml}${
    modeledFormatting.paragraphMarkPropertiesInnerXml ?? ""
  }`;
  if (paragraphMarkPropertiesInnerXml) {
    parts.push(`<w:rPr>${paragraphMarkPropertiesInnerXml}</w:rPr>`);
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
    attrs.push(`r:id="${escapeXmlAttribute(hyperlink.rId)}"`);
  }

  if (hyperlink.anchor) {
    attrs.push(`w:anchor="${escapeXmlAttribute(hyperlink.anchor)}"`);
  }

  if (hyperlink.tooltip) {
    attrs.push(`w:tooltip="${escapeXmlAttribute(hyperlink.tooltip)}"`);
  }

  if (hyperlink.target) {
    attrs.push(`w:tgtFrame="${escapeXmlAttribute(hyperlink.target)}"`);
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
    attrs.push(`w:docLocation="${escapeXmlAttribute(hyperlink.docLocation)}"`);
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
  if (child.type === "bookmarkStart") {
    return serializeBookmarkStart(child);
  }
  // Opaque markup, replayed between the same two children it was read
  // between, so a permission range or a proofing error does not leave the
  // link it was authored inside.
  return child.type === "bookmarkEnd" ? serializeBookmarkEnd(child) : child.xml;
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
  return `<w:bookmarkStart ${bookmarkRangeAttributes(bookmark).join(" ")}/>`;
}

/**
 * Serialize bookmark end (w:bookmarkEnd)
 */
function serializeBookmarkEnd(bookmark: BookmarkEnd): string {
  return `<w:bookmarkEnd ${markupRangeAttributes(bookmark).join(" ")}/>`;
}

/** Serialize a simple field without changing its authored OOXML field form. */
function serializeSimpleField(field: SimpleField): string {
  const attrs = [
    `w:instr="${escapeXmlAttribute(field.instruction)}"`,
    ...fieldStateAttributes(field),
  ];

  const contentXml = field.content
    .map((item) => {
      if (item.type === "run") {
        return serializeRun(item);
      }
      return item.type === "hyperlink" ? serializeHyperlink(item) : item.xml;
    })
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
  const beginAttrs: string[] = ['w:fldCharType="begin"', ...fieldStateAttributes(field)];
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
      `<w:r>${rPrXml}<w:instrText${spaceAttr}>${escapeXmlText(field.instruction)}</w:instrText></w:r>`,
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
    prParts.push(`<w:alias w:val="${escapeXmlAttribute(props.alias)}"/>`);
  }
  if (props.tag) {
    prParts.push(`<w:tag w:val="${escapeXmlAttribute(props.tag)}"/>`);
  }
  if (props.lock && props.lock !== "unlocked") {
    prParts.push(`<w:lock w:val="${props.lock}"/>`);
  }
  if (props.placeholder) {
    // OOXML shape: `<w:placeholder><w:docPart w:val="..."/></w:placeholder>`.
    // The placeholder identifier lives in `w:val` on the nested `w:docPart`,
    // mirroring the parse in `paragraphParser.ts`.
    prParts.push(
      `<w:placeholder><w:docPart w:val="${escapeXmlAttribute(props.placeholder)}"/></w:placeholder>`,
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
        ? ` w:fullDate="${escapeXmlAttribute(props.dateValueISO)}"`
        : "";
      const formatChild = props.dateFormat
        ? `<w:dateFormat w:val="${escapeXmlAttribute(props.dateFormat)}"/>`
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
            `<w:listItem w:displayText="${escapeXmlAttribute(i.displayText)}" w:value="${escapeXmlAttribute(i.value)}"/>`,
        )
        .join("");
      prParts.push(`<w:dropDownList>${items}</w:dropDownList>`);
      break;
    }
    case "comboBox": {
      const items = (props.listItems ?? [])
        .map(
          (i) =>
            `<w:listItem w:displayText="${escapeXmlAttribute(i.displayText)}" w:value="${escapeXmlAttribute(i.value)}"/>`,
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
        case "preservedInline":
          return item.xml;
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
  return `<w:${tag} ${moveBookmarkAttributes(marker).join(" ")}/>`;
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
    if (item.type === "mathEquation") {
      return item.ommlXml;
    }
    // Inside the wrapper, where the source put it: markup lifted out of a
    // `w:ins` is markup the reviewer no longer accepts or rejects with the
    // change.
    if (item.type === "preservedInline") {
      return item.xml;
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
 * A `commentRangeEnd` writes an end and nothing else. Where the comment's
 * reference run goes is authored data the model carries as its own item, and
 * `completeCommentReferences` fills it in for a model that arrived without
 * one; a guess made here from one paragraph cannot see the rest of the story.
 */
function serializeParagraphContent(content: ParagraphContent): string {
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
      return `<w:commentRangeStart ${markupRangeAttributes(content).join(" ")}/>`;
    case "commentRangeEnd":
      return `<w:commentRangeEnd ${markupRangeAttributes(content).join(" ")}/>`;
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
      return `<w:moveFromRangeEnd ${markupRangeAttributes(content).join(" ")}/>`;
    case "moveToRangeStart":
      return serializeMoveRangeStart("moveToRangeStart", content as MoveToRangeStart);
    case "moveToRangeEnd":
      return `<w:moveToRangeEnd ${markupRangeAttributes(content).join(" ")}/>`;
    case "bidiWrapper": {
      // `w:dir` is the embedding and `w:bdo` the override; the schema gives
      // them the same content model, which is paragraph content, so the
      // children go back through this function.
      const tag = content.control === "override" ? "bdo" : "dir";
      const value =
        content.direction === undefined ? "" : ` w:val="${escapeXmlAttribute(content.direction)}"`;
      const inner = content.content.map((child) => serializeParagraphContent(child)).join("");
      return `<w:${tag}${value}>${inner}</w:${tag}>`;
    }
    case "mathEquation":
      // Round-trip the raw OMML XML directly
      return content.ommlXml || "";
    // Opaque markup, replayed where the source put it.
    case "preservedInline":
      return content.xml;
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
    attrs.push(`w14:paraId="${escapeXmlAttribute(paragraph.paraId)}"`);
  }
  if (paragraph.textId) {
    attrs.push(`w14:textId="${escapeXmlAttribute(paragraph.textId)}"`);
  }
  if (paragraph.reviewCarrier) {
    attrs.push(`folio:reviewCarrier="${paragraph.reviewCarrier}"`);
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

  // Add paragraph content
  let pendingRenderedPageBreak = paragraph.renderedPageBreakBefore === true;
  for (const content of paragraph.content) {
    let contentXml = serializeParagraphContent(content);
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
