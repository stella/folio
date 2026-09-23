/**
 * Paragraph Parser - Parse paragraphs (w:p) with complete formatting
 *
 * A paragraph is the fundamental block-level element containing text runs,
 * hyperlinks, bookmarks, and fields.
 *
 * OOXML Reference:
 * - Paragraph: w:p
 * - Paragraph properties: w:pPr
 * - Content: runs, hyperlinks, bookmarks, fields
 */

import type {
  Paragraph,
  ParagraphContent,
  ParagraphFormatting,
  Run,
  Hyperlink,
  BookmarkStart,
  BookmarkEnd,
  SimpleField,
  ComplexField,
  TextFormatting,
  Theme,
  RelationshipMap,
  MediaFile,
  InlineSdt,
  SdtProperties,
  ParagraphMarkChange,
  ParagraphPropertyChange,
  TrackedChangeInfo,
  TrackedRunChange,
  MathEquation,
  InlineWrapper,
  RunContent,
} from "../types/document";
import { PARAGRAPH_MARK_CHANGE_KINDS, REVIEW_CARRIERS } from "@stll/docx-core/model";
import { panic } from "better-result";
import { isValidHexId } from "../utils/hexId";
import { attributeRemainder } from "./attributeRemainder";
import { paraIdAttribute, textIdAttribute } from "./paraIdAttribute";
import { paraIdInRange } from "./paraIdRangeNormalization";
import { assignParagraphPropertySource } from "./paragraphPropertySource";
import {
  parseBookmarkStart as parseBookmarkStartFromModule,
  parseBookmarkEnd as parseBookmarkEndFromModule,
} from "./bookmarkParser";
import { parseMarkupRangeMarker, parseMoveBookmarkMarker } from "./markupRangeMarker";
import { parseFieldType } from "./fieldParser";
import { type FieldState, fieldStateOf, parseFieldState } from "./fieldState";
import {
  HYPERLINK_CHILD_HANDLERS,
  type HyperlinkChildContext,
  parseHyperlink as parseHyperlinkFromModule,
  parseHyperlinkShell,
} from "./hyperlinkParser";
import {
  counterFormatOf,
  markerAlignmentForLevel,
  markerFormattingFromLevel,
  numberingLevelHasMarkerSlot,
} from "./numberingParser";
import type { NumberingMap } from "./numberingParser";
import {
  mergeParagraphNumbering,
  paragraphNumberingReferenceId,
  resolveParagraphNumbering,
} from "./numberingReference";
import { parseParagraphProperties } from "./paragraphProperties";
import {
  CAPTURE,
  type ChildHandlers,
  type ChildReader,
  dispatchChildrenWithContext,
  ownedElsewhere,
  transitionalNamespaceOf,
  withPreservedChildren,
} from "./containerChildren";
import {
  isHyperlinkContent,
  isInlineSdtContent,
  isTrackedChangeWrapperChild,
} from "./inlineWrapperContent";
import { inlineWrapperOf } from "./inlineWrapperParser";
import type { InlineWrapperElement } from "./inlineWrapperParser";
import {
  preservedInlineCapture,
  preserveInlineChild,
  preserveRunChild,
} from "./preservedRunContent";
import { type PreviewLedger, standalonePreviewLedger } from "./previewBudget";
import { consolidateParagraphContent } from "./runConsolidator";
import { parseRun } from "./runParser";
import { runHoldsPayload } from "./runPayload";
import { isVmlPictParsedByRunParser } from "./vmlImageParser";
import { captureSdtSiblingMarkers, parseSdtProperties } from "./sdtProperties";
import { parseSectionProperties } from "./sectionParser";
import type { StyleMap } from "./styleParser";
import { captureVerbatimXml } from "./verbatimCapture";
import {
  cloneElement,
  findChild,
  findChildByNamespaceUri,
  findChildrenByNamespaceUri,
  findWordprocessingChild,
  getAttributeByNamespaceUri,
  getAttribute,
  getChildElements,
  getLocalName,
  getNamespaceUri,
  mergeXmlnsDeclarations,
  parseBooleanElement,
  parseNumericAttribute,
  selectAlternateContentBranch,
  WORDPROCESSINGML_NAMESPACE_URIS,
} from "./xmlParser";
import type { XmlElement } from "./xmlParser";
import { hasAttributeAnySpelling } from "./strictNames";
import { scanRunForTextBoxDrawings } from "./textBoxParser";
import { parsePropertyChangeInfo, parseTrackedChangeInfo } from "./trackedChangeInfo";

const FOLIO_REVIEW_HISTORY_NAMESPACE = "urn:stella:folio:review-history:1";
const FOLIO_REVIEW_HISTORY_NAMESPACES: ReadonlySet<string> = new Set([
  FOLIO_REVIEW_HISTORY_NAMESPACE,
]);

/**
 * Extract plain text from a math element (recursive text content extraction)
 */
function extractPlainText(runs: Run[]): string {
  let out = "";
  for (const run of runs) {
    for (const c of run.content) {
      if (c.type === "text") {
        out += c.text;
      }
    }
  }
  return out;
}

function extractMathText(el: XmlElement): string {
  let text = "";
  if (el.type === "text" && typeof el.text === "string") {
    return el.text;
  }
  if (el.elements) {
    for (const child of el.elements) {
      // m:t elements contain the actual math text
      const childName = child.name?.replace(/^.*:/u, "") ?? "";
      if (childName === "t" && child.elements) {
        for (const t of child.elements) {
          if (t.type === "text" && typeof t.text === "string") {
            text += t.text;
          }
        }
      } else {
        text += extractMathText(child);
      }
    }
  }
  return text;
}
/**
 * Capture the authored paragraph properties separately from structural and
 * revision children, which have their own model fields and lifecycles.
 */
const captureParagraphPropertySource = (pPr: XmlElement): string =>
  captureVerbatimXml(
    cloneElement(pPr, {
      elements: (pPr.elements ?? []).flatMap((child) => {
        if (child.type !== "element") {
          return [child];
        }
        const localName = getLocalName(child.name);
        if (
          WORDPROCESSINGML_NAMESPACE_URIS.has(child.namespaceUri ?? "") &&
          (localName === "sectPr" || localName === "pPrChange")
        ) {
          return [];
        }
        if (localName !== "rPr" || !WORDPROCESSINGML_NAMESPACE_URIS.has(child.namespaceUri ?? "")) {
          return [child];
        }
        return [
          cloneElement(child, {
            elements: (child.elements ?? []).filter((runProperty) => {
              if (runProperty.type !== "element") {
                return true;
              }
              const runPropertyName = getLocalName(runProperty.name);
              return (
                !WORDPROCESSINGML_NAMESPACE_URIS.has(runProperty.namespaceUri ?? "") ||
                !PARAGRAPH_MARK_CHANGE_KINDS.some((kind) => kind === runPropertyName)
              );
            }),
          }),
        ];
      }),
    }),
  );

// ============================================================================
// PARAGRAPH CONTENT PARSERS
// ============================================================================

const RENDERED_BREAK_INLINE_WRAPPERS = new Set([
  "hyperlink",
  "smartTag",
  "sdt",
  "sdtContent",
  "fldSimple",
  "customXml",
  "ins",
  "del",
  "moveFrom",
  "moveTo",
]);

const RENDERED_BREAK_NON_CONTENT_MARKERS = new Set([
  "pPr",
  "proofErr",
  "bookmarkStart",
  "bookmarkEnd",
  "commentRangeStart",
  "commentRangeEnd",
  "commentReference",
  "permStart",
  "permEnd",
  "rsidR",
  "sdtPr",
  "sdtEndPr",
  "smartTagPr",
]);

const RENDERED_BREAK_VISIBLE_RUN_CONTENT = new Set([
  "AlternateContent",
  "t",
  "tab",
  "br",
  "cr",
  "sym",
  "drawing",
  "pict",
  "object",
  "softHyphen",
  "noBreakHyphen",
  "fldChar",
  "instrText",
  "pgNum",
  "separator",
  "continuationSeparator",
  "footnoteRef",
  "endnoteRef",
  "footnoteReference",
  "endnoteReference",
  "ptab",
  "monthShort",
  "monthLong",
  "yearShort",
  "yearLong",
  "dayShort",
  "dayLong",
]);

function paragraphStartsWithRenderedPageBreak(node: XmlElement): boolean {
  type VisitResult = "forced" | "visible" | "continue";
  let sawRenderedPageBreak = false;

  const visit = (element: XmlElement): VisitResult => {
    for (const child of getChildElements(element)) {
      const childName = getLocalName(child.name);
      if (RENDERED_BREAK_NON_CONTENT_MARKERS.has(childName)) {
        continue;
      }
      if (childName === "lastRenderedPageBreak") {
        sawRenderedPageBreak = true;
        continue;
      }
      if (childName === "r") {
        for (const runChild of getChildElements(child)) {
          const runChildName = getLocalName(runChild.name);
          if (runChildName === "rPr") {
            continue;
          }
          if (runChildName === "lastRenderedPageBreak") {
            sawRenderedPageBreak = true;
            continue;
          }
          if (runChildName === "br" && getAttribute(runChild, "w", "type") === "page") {
            return "forced";
          }
          if (RENDERED_BREAK_VISIBLE_RUN_CONTENT.has(runChildName)) {
            return "visible";
          }
        }
        continue;
      }
      if (RENDERED_BREAK_INLINE_WRAPPERS.has(childName)) {
        const result = visit(child);
        if (result !== "continue") {
          return result;
        }
        continue;
      }
      return "continue";
    }
    return "continue";
  };

  const outcome = visit(node);
  if (outcome === "forced") {
    // A hard break is structural content, not Word's cached pagination metadata.
    // Treating it as both makes a following w:lastRenderedPageBreak advance the
    // advisory page target twice.
    return sawRenderedPageBreak;
  }
  return outcome === "visible" && sawRenderedPageBreak;
}

type TrackedChangeParseContext = "default" | "deletion";
type TrackedChangeWrapperType = "insertion" | "deletion" | "moveFrom" | "moveTo";

function replaceLocalName(name: string | undefined, localName: string): string {
  if (!name) {
    return `w:${localName}`;
  }
  const colonIndex = name.indexOf(":");
  if (colonIndex === -1) {
    return localName;
  }
  return `${name.slice(0, colonIndex + 1)}${localName}`;
}

function normalizeDeletionContentElement(node: XmlElement): XmlElement {
  if (node.type !== "element") {
    return node;
  }

  const localName = getLocalName(node.name);
  let mappedName = node.name;

  if (localName === "delText") {
    mappedName = replaceLocalName(node.name, "t");
  } else if (localName === "delInstrText") {
    mappedName = replaceLocalName(node.name, "instrText");
  }

  return cloneElement(node, {
    ...(mappedName !== undefined ? { name: mappedName } : {}),
    ...(node.elements ? { elements: node.elements.map(normalizeDeletionContentElement) } : {}),
  });
}

export function parseParagraphPropertyChanges(
  pPr: XmlElement | null,
  theme: Theme | null,
  currentFormatting: ParagraphFormatting | undefined,
): ParagraphPropertyChange[] | undefined {
  if (!pPr) {
    return undefined;
  }

  const changes = findChildrenByNamespaceUri(pPr, WORDPROCESSINGML_NAMESPACE_URIS, "pPrChange").map(
    (changeElement): ParagraphPropertyChange => {
      const previousPPr = findChildByNamespaceUri(
        changeElement,
        WORDPROCESSINGML_NAMESPACE_URIS,
        "pPr",
      );
      const previousFormatting = parseParagraphProperties(previousPPr, theme);
      const change: ParagraphPropertyChange = {
        type: "paragraphPropertyChange",
        info: parsePropertyChangeInfo(changeElement),
      };
      if (previousFormatting !== undefined) {
        change.previousFormatting = previousFormatting;
      }
      if (currentFormatting !== undefined) {
        change.currentFormatting = currentFormatting;
      }
      return change;
    },
  );

  return changes.length > 0 ? changes : undefined;
}

/**
 * Parse `<w:pPr><w:rPr><w:ins/>` or `<w:del/>` — the OOXML paragraph-mark
 * tracked-change marker (ECMA-376 §17.13.5). Returns the first marker
 * encountered, or undefined when none is present. `ins` and `del` are
 * mutually exclusive in valid documents.
 */
function parseParagraphMarkChange(pPr: XmlElement | null): ParagraphMarkChange | undefined {
  if (!pPr) {
    return undefined;
  }
  const rPr = findChildByNamespaceUri(pPr, WORDPROCESSINGML_NAMESPACE_URIS, "rPr");
  if (!rPr) {
    return undefined;
  }
  // Driven by the kinds themselves, so a kind the model gains cannot be one
  // the parser silently drops. `moveFrom` / `moveTo` come first: a moved
  // paragraph's mark carries one of them INSTEAD of `del` / `ins`, never both.
  for (const kind of PARAGRAPH_MARK_CHANGE_KINDS) {
    const element = findChildByNamespaceUri(rPr, WORDPROCESSINGML_NAMESPACE_URIS, kind);
    if (element) {
      return { kind, info: parseTrackedChangeInfo(element) };
    }
  }
  return undefined;
}

type PushTrackedChangeWrapperParams = {
  contents: ParagraphContent[];
  type: TrackedChangeWrapperType;
  info: TrackedChangeInfo;
  content: readonly TrackedRunChange["content"][number][];
  preserveEmpty?: boolean;
};

function pushTrackedChangeWrapper({
  contents,
  type,
  info,
  content,
  preserveEmpty = false,
}: PushTrackedChangeWrapperParams): void {
  if (content.length === 0 && !preserveEmpty) {
    return;
  }

  if (type === "insertion") {
    contents.push({ type: "insertion", info, content: [...content] });
    return;
  }

  if (type === "deletion") {
    contents.push({ type: "deletion", info, content: [...content] });
    return;
  }

  if (type === "moveFrom") {
    contents.push({ type: "moveFrom", info, content: [...content] });
    return;
  }

  contents.push({ type: "moveTo", info, content: [...content] });
}

type PushTrackedChangeSegmentsParams = {
  contents: ParagraphContent[];
  type: TrackedChangeWrapperType;
  info: TrackedChangeInfo;
  parsedContent: readonly ParagraphContent[];
};

function pushTrackedChangeSegments({
  contents,
  type,
  info,
  parsedContent,
}: PushTrackedChangeSegmentsParams): void {
  if (!parsedContent.some(isTrackedChangeWrapperChild)) {
    pushTrackedChangeWrapper({
      contents,
      type,
      info,
      content: [],
      preserveEmpty: true,
    });
    contents.push(...parsedContent);
    return;
  }

  const segment: TrackedRunChange["content"] = [];

  for (const content of parsedContent) {
    if (isTrackedChangeWrapperChild(content)) {
      segment.push(content);
      continue;
    }

    pushTrackedChangeWrapper({ contents, type, info, content: segment });
    segment.length = 0;
    contents.push(content);
  }

  pushTrackedChangeWrapper({ contents, type, info, content: segment });
}

type PushInlineSdtSegmentsParams = {
  contents: ParagraphContent[];
  properties: SdtProperties;
  parsedContent: readonly ParagraphContent[];
};

function pushInlineSdtSegments({
  contents,
  properties,
  parsedContent,
}: PushInlineSdtSegmentsParams): void {
  if (!parsedContent.some(isInlineSdtContent)) {
    contents.push({
      type: "inlineSdt",
      properties,
      content: [],
    });
    contents.push(...parsedContent);
    return;
  }

  const segment: InlineSdt["content"] = [];

  const pushSegment = (): void => {
    if (segment.length === 0) {
      return;
    }

    contents.push({
      type: "inlineSdt",
      properties,
      content: [...segment],
    });
    segment.length = 0;
  };

  for (const content of parsedContent) {
    if (isInlineSdtContent(content)) {
      segment.push(content);
      continue;
    }

    pushSegment();
    contents.push(content);
  }

  pushSegment();
}

/**
 * Parse hyperlink element (w:hyperlink)
 *
 * Delegates to hyperlinkParser module which resolves URLs via relationships.
 */
function parseHyperlink(
  node: XmlElement,
  rels: RelationshipMap | null,
  styles: StyleMap | null,
  theme: Theme | null,
  media: Map<string, MediaFile> | null,
  rootXmlns: Record<string, string>,
  previews: PreviewLedger,
): Hyperlink {
  return parseHyperlinkFromModule(node, rels, styles, theme, media, rootXmlns, previews);
}

/**
 * The four `CT_RunTrackChange` wrappers a `w:hyperlink` may hold.
 *
 * What to *do* with each is the dispatcher's handler map; this set only
 * answers whether the link needs segmenting at all, which decides between one
 * link and a sequence of links and revisions.
 */
const HYPERLINK_REVISION_WRAPPERS: ReadonlySet<string> = new Set([
  "del",
  "ins",
  "moveFrom",
  "moveTo",
]);

const OMML_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/math";

/**
 * An OMML element as an equation, or nothing when the child is not one.
 *
 * `m:oMathPara` is the display form and `m:oMath` the inline one. The rest of
 * `m:EG_OMathMathElements` — `m:f`, `m:acc`, `m:rad` and their siblings — the
 * schema admits wherever `m:oMath` is admitted, so `<w:ins><m:f/></w:ins>` is
 * a tracked insertion of a fraction with no `m:oMath` around it. None of them
 * carries structure the editable model holds, so all travel as the markup
 * they arrived as.
 */
const mathContentOf = (child: XmlElement): MathEquation | undefined => {
  const namespace = getNamespaceUri(child);
  // A Strict package spells the maths namespace under `purl.oclc.org`, and an
  // equation is an equation in either class.
  if (namespace === undefined || transitionalNamespaceOf(namespace) !== OMML_NAMESPACE) {
    return undefined;
  }
  const equation: MathEquation = {
    type: "mathEquation",
    display: getLocalName(child.name) === "oMathPara" ? "block" : "inline",
    ommlXml: captureVerbatimXml(child),
  };
  const plainText = extractMathText(child);
  if (plainText) {
    equation.plainText = plainText;
  }
  return equation;
};

/** The link's revision-segmenting walk: the link's own, plus where it records a revision. */
type HyperlinkRevisionWalk = HyperlinkChildContext & {
  items: (Hyperlink["children"][number] | HoistedRevision)[];
  linkOver: (linkChildren: readonly Hyperlink["children"][number][]) => Hyperlink;
};

const hoistRevision =
  (wrapper: TrackedChangeWrapperType): ChildReader<HyperlinkRevisionWalk> =>
  (child, { styles, theme, rels, media, previews, inScopeXmlns, items, linkOver }) => {
    const wrapped = parseParagraphContents(
      child,
      styles,
      theme,
      null,
      rels,
      media,
      previews,
      wrapper === "deletion" || wrapper === "moveFrom" ? "deletion" : "default",
      inScopeXmlns,
    );
    // Group the runs the wrapper holds back under the link; anything else it
    // carries stays where it sits rather than being dropped.
    const content: TrackedRunChange["content"][number][] = [];
    let linked: Hyperlink["children"][number][] = [];
    const flushLinked = (): void => {
      if (linked.length > 0) {
        content.push(linkOver(linked));
        linked = [];
      }
    };
    for (const item of wrapped) {
      if (isHyperlinkContent(item)) {
        linked.push(item);
        continue;
      }
      flushLinked();
      if (isTrackedChangeWrapperChild(item)) {
        content.push(item);
      }
    }
    flushLinked();
    items.push({
      type: "hoistedRevision",
      wrapper,
      info: parseTrackedChangeInfo(child),
      content,
    });
  };

const HYPERLINK_REVISION_HOISTING_HANDLERS = {
  ...HYPERLINK_CHILD_HANDLERS,
  ins: hoistRevision("insertion"),
  del: hoistRevision("deletion"),
  moveFrom: hoistRevision("moveFrom"),
  moveTo: hoistRevision("moveTo"),
} as const satisfies ChildHandlers<"w:hyperlink", HyperlinkRevisionWalk>;

/**
 * A `w:hyperlink` as paragraph content, with any revision wrapper it holds
 * hoisted around it.
 *
 * OOXML nests `w:ins`/`w:del` INSIDE `w:hyperlink`; the model nests the
 * hyperlink inside the revision, because a revision is the unit a redline
 * reads and a link that is half deleted is two links to it. This is the exact
 * inverse of what the serializer writes, so a package survives the round trip.
 */
function parseHyperlinkParagraphContents(
  node: XmlElement,
  rels: RelationshipMap | null,
  styles: StyleMap | null,
  theme: Theme | null,
  media: Map<string, MediaFile> | null,
  rootXmlns: Record<string, string>,
  previews: PreviewLedger,
): ParagraphContent[] {
  const children = getChildElements(node);
  if (!children.some((child) => HYPERLINK_REVISION_WRAPPERS.has(getLocalName(child.name)))) {
    return [parseHyperlink(node, rels, styles, theme, media, rootXmlns, previews)];
  }

  const inScopeXmlns = mergeXmlnsDeclarations(rootXmlns, node);
  const shell = parseHyperlinkShell(node, rels);
  const linkOver = (linkChildren: readonly Hyperlink["children"][number][]): Hyperlink => ({
    ...shell,
    children: [...linkChildren],
  });

  // The walk is flat and the segmenting happens after it, so the dispatcher's
  // sink can index an undeclared child against one list rather than against
  // whichever segment happened to be open when it was read.
  const items: (Hyperlink["children"][number] | HoistedRevision)[] = [];
  const preserved = dispatchChildrenWithContext({
    element: node,
    container: "w:hyperlink",
    capturePosition: () => items.length,
    handlers: HYPERLINK_REVISION_HOISTING_HANDLERS,
    context: {
      push: (child) => {
        items.push(child);
      },
      styles,
      theme,
      rels,
      media,
      previews,
      inScopeXmlns,
      items,
      linkOver,
    },
  });

  const contents: ParagraphContent[] = [];
  let plain: Hyperlink["children"][number][] = [];
  const flushPlain = (): void => {
    if (plain.length > 0) {
      contents.push(linkOver(plain));
      plain = [];
    }
  };

  for (const item of withPreservedChildren(items, preserved, preservedInlineCapture)) {
    if (item.type !== "hoistedRevision") {
      plain.push(item);
      continue;
    }
    flushPlain();
    pushTrackedChangeWrapper({
      contents,
      type: item.wrapper,
      info: item.info,
      content: item.content,
      preserveEmpty: true,
    });
  }
  flushPlain();

  return contents;
}

/**
 * A `CT_RunTrackChange` read out of a `w:hyperlink`, before it is hoisted
 * around the link.
 *
 * OOXML nests the revision inside the link and the model nests the link
 * inside the revision, so the two cannot be built in one pass: the walk
 * records the revision in source order and the segmenting loop turns it into
 * the wrapper.
 */
type HoistedRevision = {
  type: "hoistedRevision";
  wrapper: TrackedChangeWrapperType;
  info: TrackedChangeInfo;
  content: readonly TrackedRunChange["content"][number][];
};

/**
 * Parse bookmark start (w:bookmarkStart)
 * Delegates to bookmarkParser module.
 */
function parseBookmarkStart(node: XmlElement): BookmarkStart {
  return parseBookmarkStartFromModule(node);
}

/**
 * Parse bookmark end (w:bookmarkEnd)
 * Delegates to bookmarkParser module.
 */
function parseBookmarkEnd(node: XmlElement): BookmarkEnd {
  return parseBookmarkEndFromModule(node);
}

/**
 * Parse simple field (w:fldSimple)
 */
function parseSimpleField(
  node: XmlElement,
  styles: StyleMap | null,
  theme: Theme | null,
  rels: RelationshipMap | null,
  media: Map<string, MediaFile> | null,
  rootXmlns: Record<string, string>,
  previews: PreviewLedger,
): SimpleField {
  const instruction = getAttribute(node, "w", "instr") ?? "";
  const fieldType = parseFieldType(instruction);

  const field: SimpleField = {
    type: "simpleField",
    instruction,
    fieldType,
    content: [],
    ...parseFieldState(node),
  };

  // Parse display content without changing its authored field form.
  //
  // `CT_SimpleField` is `EG_PContent` plus `w:fldData`, so a field's cached
  // result may hold everything a paragraph may: a bookmark around the result,
  // a proofing error, a nested revision. folio models the run and the link;
  // the rest is markup it carries at its source position rather than markup
  // it drops.
  const inScopeXmlns = mergeXmlnsDeclarations(rootXmlns, node);
  const content: SimpleField["content"] = [];
  const preserved = dispatchChildrenWithContext({
    element: node,
    container: "w:fldSimple",
    capturePosition: () => content.length,
    handlers: SIMPLE_FIELD_CHILD_HANDLERS,
    context: {
      push: (child) => {
        content.push(child);
      },
      styles,
      theme,
      rels,
      media,
      previews,
      inScopeXmlns,
    },
  });
  field.content = withPreservedChildren(content, preserved, preservedInlineCapture);

  return field;
}

/** What {@link SIMPLE_FIELD_CHILD_HANDLERS} needs to read one child of a field. */
type SimpleFieldChildContext = {
  /** Where a parsed or captured child lands, in source order. */
  push: (child: SimpleField["content"][number]) => void;
  styles: StyleMap | null;
  theme: Theme | null;
  rels: RelationshipMap | null;
  media: Map<string, MediaFile> | null;
  previews: PreviewLedger;
  inScopeXmlns: Record<string, string>;
};

/** A transparent wrapper the field holds, read as the wrapper it is. */
const fieldInlineWrapper =
  (element: InlineWrapperElement): ChildReader<SimpleFieldChildContext> =>
  (child, context) => {
    context.push(parseFieldInlineWrapper(element, child, context));
  };

/**
 * What a `w:fldSimple` does with every child its content model declares.
 *
 * Two callers read this one map: {@link parseSimpleField}, and
 * {@link parseFieldInlineWrapper} for a transparent wrapper the field holds —
 * `CT_BdoContentRun` and its three siblings are the same `EG_PContent` the
 * field is, so the decision per child is the field's own.
 */
const SIMPLE_FIELD_CHILD_HANDLERS = {
  r: (child, { push, styles, theme, rels, media, previews, inScopeXmlns }) => {
    push(parseRun(child, styles, theme, rels, media, inScopeXmlns, previews));
  },
  hyperlink: (child, { push, styles, theme, rels, media, previews, inScopeXmlns }) => {
    push(parseHyperlink(child, rels, styles, theme, media, inScopeXmlns, previews));
  },

  // The transparent wrappers, read as the wrappers they are: a cached field
  // result written inside a `w:dir` keeps its runs editable.
  bdo: fieldInlineWrapper("bdo"),
  dir: fieldInlineWrapper("dir"),
  customXml: fieldInlineWrapper("customXml"),
  smartTag: fieldInlineWrapper("smartTag"),

  bookmarkEnd: CAPTURE,
  bookmarkStart: CAPTURE,
  commentRangeEnd: CAPTURE,
  commentRangeStart: CAPTURE,
  customXmlDelRangeEnd: CAPTURE,
  customXmlDelRangeStart: CAPTURE,
  customXmlInsRangeEnd: CAPTURE,
  customXmlInsRangeStart: CAPTURE,
  customXmlMoveFromRangeEnd: CAPTURE,
  customXmlMoveFromRangeStart: CAPTURE,
  customXmlMoveToRangeEnd: CAPTURE,
  customXmlMoveToRangeStart: CAPTURE,
  del: CAPTURE,
  // The field's own custom data (`CT_Text`), meaningful only to the
  // producer that wrote it, so it travels as the bytes it arrived as.
  fldData: CAPTURE,
  fldSimple: (child, { push }) => {
    push(preserveInlineChild(child));
  },
  ins: CAPTURE,
  moveFrom: CAPTURE,
  moveFromRangeEnd: CAPTURE,
  moveFromRangeStart: CAPTURE,
  moveTo: CAPTURE,
  moveToRangeEnd: CAPTURE,
  moveToRangeStart: CAPTURE,
  permEnd: CAPTURE,
  permStart: CAPTURE,
  proofErr: CAPTURE,
  sdt: CAPTURE,
  subDoc: CAPTURE,
} as const satisfies ChildHandlers<"w:fldSimple", SimpleFieldChildContext>;

/**
 * A transparent wrapper a simple field holds, with the content the field holds.
 *
 * The mirror of `parseLinkedInlineWrapper`: the wrapper's declared children are
 * run-level content, the decision per child is the container's own map, and the
 * wrapper's properties bag is read by `inlineWrapperOf` rather than captured a
 * second time by the sink.
 */
const parseFieldInlineWrapper = (
  element: InlineWrapperElement,
  node: XmlElement,
  context: SimpleFieldChildContext,
): InlineWrapper => {
  const inScopeXmlns = mergeXmlnsDeclarations(context.inScopeXmlns, node);
  const content: SimpleField["content"] = [];
  const preserved = dispatchChildrenWithContext({
    element: node,
    container: "run-level-content",
    capturePosition: () => content.length,
    handlers: FIELD_INLINE_WRAPPER_HANDLERS,
    context: {
      ...context,
      inScopeXmlns,
      push: (child) => {
        content.push(child);
      },
    },
  });
  return inlineWrapperOf(
    element,
    node,
    withPreservedChildren(content, preserved, preservedInlineCapture),
  );
};

/**
 * Whether a run is worth keeping once it has been parsed.
 *
 * This asks the model. A keep rule that reads the source element and a writer
 * that reads the model can only agree while the model is complete, and the
 * disagreement is a two-save oscillation rather than a loss: the first save
 * writes a run whose payload the model never held, the next parse drops that
 * run, and the second save differs from the first. Every unmodelled run child
 * now reaches `content` as a preserved capture, so {@link runHoldsPayload},
 * the one predicate the consolidator and the serializer ask too, answers the
 * question for all of them.
 *
 * The one exception is not an unmodelled child but an unfinished model: a
 * text box is claimed by `enrichParagraphTextBoxes`, a second pass over the
 * same paragraph, so its run is legitimately empty here and dropping it would
 * take the text box with it. `scanRunForTextBoxDrawings` is the pass's own
 * reader, called rather than restated so the two cannot disagree about which
 * runs it will claim.
 */
type HasRunPayloadOptions = {
  run: Run;
  runElement: XmlElement;
  rels: RelationshipMap | null;
  media: Map<string, MediaFile> | null;
};

const hasRunPayload = ({ run, runElement, rels, media }: HasRunPayloadOptions): boolean => {
  if (runHoldsPayload(run)) {
    return true;
  }
  const { textBoxDrawings, vmlTextBoxes } = scanRunForTextBoxDrawings({
    xmlRun: runElement,
    claimedByRunParser: (pictElement) => isVmlPictParsedByRunParser(pictElement, rels, media),
  });
  return textBoxDrawings.length > 0 || vmlTextBoxes.length > 0;
};

/**
 * A field character that belongs to no field, kept as the markup it was.
 *
 * `FieldCharContent` is half of a `ComplexField`: `convertField` reads it off
 * the assembled field, and one that reaches the editor's inline converter on
 * its own is dropped. So a `w:fldChar` whose field this paragraph never closed,
 * or one with no `begin` before it, loses the editor round trip — and the
 * model holds nothing of its `w:ffData` either, so the save loses a legacy form
 * field's name, macros, help text and checkbox state as well.
 *
 * `preservedXml` is the run's own capture member: the editor carries it as an
 * opaque atom and the serializer replays the source bytes, `w:ffData` included.
 * The two lists line up because `parseRunContents` walks the source children in
 * order, so the nth `w:fldChar` child is the nth `fieldChar` content.
 *
 * `w:instrText` is left alone. It is fully modelled and fully serialized; what
 * the editor does with an orphan is a separate question from what the container
 * contract asks here, and answering it with bytes would trade a model for one.
 */
const withOrphanFieldCharsPreserved = (run: Run, runElement: XmlElement): Run => {
  const captured = getChildElements(runElement)
    .filter((child) => getLocalName(child.name) === "fldChar")
    .map(preserveRunChild);
  if (captured.length === 0) {
    return run;
  }
  let next = 0;
  const content = run.content.map((item) => {
    if (item.type !== "fieldChar") {
      return item;
    }
    const replacement = captured[next];
    next += 1;
    return replacement ?? item;
  });
  return { ...run, content };
};

const LEGACY_FORM_CHECKBOX_GLYPHS = {
  checked: "☒",
  unchecked: "☐",
} as const;
const LEGACY_FORM_CHECKBOX_INSTRUCTION = "FORMCHECKBOX";

type LegacyFormCheckboxDisplay = {
  text: string;
  fontSize?: number;
};

function getLegacyFormCheckboxDisplay(
  runElement: XmlElement,
): LegacyFormCheckboxDisplay | undefined {
  const fieldChar = findChild(runElement, "w", "fldChar");
  const fieldData = fieldChar ? findChild(fieldChar, "w", "ffData") : null;
  const checkBox = fieldData ? findChild(fieldData, "w", "checkBox") : null;
  if (!checkBox) {
    return undefined;
  }

  const checked = findChild(checkBox, "w", "checked");
  const defaultChecked = findChild(checkBox, "w", "default");
  let isChecked = defaultChecked ? parseBooleanElement(defaultChecked) : false;
  if (checked) {
    isChecked = parseBooleanElement(checked);
  }
  const explicitSize = parseNumericAttribute(findChild(checkBox, "w", "size"), "w", "val");
  const text = isChecked
    ? LEGACY_FORM_CHECKBOX_GLYPHS.checked
    : LEGACY_FORM_CHECKBOX_GLYPHS.unchecked;
  if (explicitSize === undefined) {
    return { text };
  }

  return {
    text,
    fontSize: explicitSize,
  };
}

function createLegacyFormCheckboxResultRun(
  display: LegacyFormCheckboxDisplay,
  inheritedFormatting: TextFormatting | undefined,
): Run {
  const run: Run = {
    type: "run",
    content: [{ type: "text", text: display.text }],
  };
  const hasInheritedFormatting =
    inheritedFormatting !== undefined && Object.keys(inheritedFormatting).length > 0;

  if (display.fontSize !== undefined) {
    run.formatting = hasInheritedFormatting
      ? { ...inheritedFormatting, fontSize: display.fontSize }
      : { fontSize: display.fontSize };
  } else if (hasInheritedFormatting) {
    run.formatting = inheritedFormatting;
  }

  return run;
}

function isLegacyFormCheckboxInstruction(instruction: string): boolean {
  const instructionName = instruction.trim().split(/\s+/u).at(0)?.toUpperCase();
  return instructionName === LEGACY_FORM_CHECKBOX_INSTRUCTION;
}

/**
 * The paragraph's own properties, read from the element by `parseParagraph`.
 *
 * The entry is in the inline handler map because the map covers every inline
 * container, not because this walk reads it. Declared at module scope so the
 * claim is registered when the module loads.
 */
const PARAGRAPH_PROPERTIES_OWNER = ownedElsewhere({
  container: "run-level-content",
  child: "pPr",
  reader: "paragraphProperties#parseParagraphProperties",
});

const SMART_TAG_PROPERTIES_OWNER = ownedElsewhere({
  container: "run-level-content",
  child: "smartTagPr",
  reader: "inlineWrapperParser#inlineWrapperOf",
});

const CUSTOM_XML_PROPERTIES_OWNER = ownedElsewhere({
  container: "run-level-content",
  child: "customXmlPr",
  reader: "inlineWrapperParser#inlineWrapperOf",
});

/** The field's own map over a transparent wrapper it holds; see `parseFieldInlineWrapper`. */
const FIELD_INLINE_WRAPPER_HANDLERS = {
  ...SIMPLE_FIELD_CHILD_HANDLERS,
  customXmlPr: CUSTOM_XML_PROPERTIES_OWNER,
  smartTagPr: SMART_TAG_PROPERTIES_OWNER,
  // Not a child of any of the four wrappers; the declared set is shared
  // with `w:p`, where the paragraph reads it off the element.
  pPr: CAPTURE,
} as const satisfies ChildHandlers<"run-level-content", SimpleFieldChildContext>;

/**
 * What one paragraph-content walk reads with and writes into.
 *
 * `scan` is the complex-field state machine: the `w:r` handler advances it run
 * by run, and the paragraph reads what is left open once the walk is done.
 */
type ParagraphContentsWalk = {
  styles: StyleMap | null;
  theme: Theme | null;
  rels: RelationshipMap | null;
  media: Map<string, MediaFile> | null;
  previews: PreviewLedger;
  trackedContext: TrackedChangeParseContext;
  inScopeXmlns: Record<string, string>;
  contents: ParagraphContent[];
  scan: ComplexFieldScan;
};

/** The complex field a paragraph walk has open, if any. */
type ComplexFieldScan = {
  inComplexField: boolean;
  complexFieldInstr: string;
  complexFieldCodeRuns: Run[];
  complexFieldResultRuns: Run[];
  // Every run read since the `begin`, whole and in source order. A field the
  // paragraph never closes is not a field, and the runs it swallowed — their
  // text, their field characters, their `w:instrText` — are ordinary content
  // that has to come back out. Keeping the runs themselves rather than the
  // code/result split is what puts the `begin` back too, so a field closed in
  // a later paragraph still has the character that opens it.
  complexFieldOpenRuns: Run[];
  afterSeparator: boolean;
  complexFieldState: FieldState;
  complexFieldFallbackDisplay: LegacyFormCheckboxDisplay | undefined;
  // Run formatting (w:rPr) carried on the field's structural runs, used as a
  // fallback when the field has no separate result run (eigenpal/docx-editor#909).
  complexFieldFormatting: TextFormatting | undefined;
};

/**
 * A transparent wrapper at paragraph level. All four hold ordinary inline
 * content and say something about it rather than about the text, so the
 * recursion is this same walk and `inlineWrapperOf` reads what the element
 * adds.
 */
const paragraphInlineWrapper =
  (element: InlineWrapperElement): ChildReader<ParagraphContentsWalk> =>
  (child, { styles, theme, rels, media, previews, trackedContext, inScopeXmlns, contents }) => {
    contents.push(
      inlineWrapperOf(
        element,
        child,
        parseParagraphContents(
          child,
          styles,
          theme,
          null,
          rels,
          media,
          previews,
          trackedContext,
          mergeXmlnsDeclarations(inScopeXmlns, child),
        ),
      ),
    );
  };

const PARAGRAPH_CONTENT_UNDECLARED = {
  // `mc:AlternateContent` is markup compatibility, legal wherever its
  // fallback is. folio selects a branch and reads it; capturing the
  // wrapper whole would keep the bytes and lose every run inside it.
  AlternateContent: (
    child,
    { contents, styles, theme, rels, media, previews, trackedContext, inScopeXmlns },
  ) => {
    const selectedBranch = selectAlternateContentBranch(child);
    if (selectedBranch) {
      contents.push(
        ...parseParagraphContents(
          selectedBranch,
          styles,
          theme,
          null,
          rels,
          media,
          previews,
          trackedContext,
          mergeXmlnsDeclarations(inScopeXmlns, child),
        ),
      );
    }
  },
} as const satisfies Record<string, ChildReader<ParagraphContentsWalk>>;

const PARAGRAPH_CONTENT_UNDECLARED_NAMESPACES = {
  // A bare OMML element is inline content in its own right: every group
  // that admits `m:oMath` also admits `m:EG_OMathMathElements`, so
  // `<w:ins><m:f/></w:ins>` is a tracked insertion of a fraction with no
  // `m:oMath` around it. Reading only the wrapper left the revision in
  // the document with its content gone, which is a reviewer accepting an
  // edit that is no longer there.
  [OMML_NAMESPACE]: (child, { contents }) => {
    const equation = mathContentOf(child);
    if (equation !== undefined) {
      contents.push(equation);
    }
  },
} as const satisfies Record<string, ChildReader<ParagraphContentsWalk>>;

const PARAGRAPH_CONTENT_HANDLERS = {
  r: (
    child,
    { contents, styles, theme, rels, media, previews, trackedContext, inScopeXmlns, scan },
  ) => {
    // Check for field characters in this run
    const runElement =
      trackedContext === "deletion" ? normalizeDeletionContentElement(child) : child;
    const run = parseRun(runElement, styles, theme, rels, media, inScopeXmlns, previews);
    const commentReferenceId = getCommentReferenceId(runElement);

    // Look for field characters
    let hasFieldBegin = false;
    let beginFieldState: FieldState = {};
    const beginFallbackDisplay = getLegacyFormCheckboxDisplay(runElement);
    let hasFieldSeparate = false;
    let hasFieldEnd = false;
    let endOriginalValue: string | undefined;
    let instrText = "";

    for (const content of run.content) {
      if (content.type === "fieldChar") {
        if (content.charType === "begin") {
          hasFieldBegin = true;
          beginFieldState = fieldStateOf(content);
        } else if (content.charType === "separate") {
          hasFieldSeparate = true;
        } else {
          hasFieldEnd = true;
          if (content.originalValue !== undefined) {
            endOriginalValue = content.originalValue;
          }
        }
      } else if (content.type === "instrText") {
        instrText += content.text;
      }
    }

    if (hasFieldBegin) {
      // Nested complex field. Word allows fields inside the result region
      // of another field (e.g. PAGEREF inside a TOC entry). We don't model
      // nesting in ParagraphContent, so before resetting state for the
      // inner field, flush any result-region runs the outer field has
      // already accumulated (e.g. the TOC entry text + tab before its
      // PAGEREF) into `contents` so they don't get destroyed by the reset.
      // Pre-separator nesting (a field inside the outer's field code) is
      // exotic enough to leave to the outer's fieldCode array.
      if (scan.inComplexField && scan.afterSeparator) {
        contents.push(...scan.complexFieldResultRuns);
        scan.inComplexField = false;
      }
      scan.inComplexField = true;
      scan.afterSeparator = false;
      scan.complexFieldInstr = "";
      scan.complexFieldCodeRuns = [];
      scan.complexFieldResultRuns = [];
      scan.complexFieldOpenRuns = [];
      // `w:fldLock` / `w:dirty` live on the begin fldChar of this field.
      scan.complexFieldState = beginFieldState;
      scan.complexFieldFallbackDisplay = beginFallbackDisplay;
      // The structural run carrying `begin` often holds the field's run
      // formatting (e.g. a footer PAGE field collapsed into one run).
      scan.complexFieldFormatting = run.formatting;
    }

    if (scan.inComplexField) {
      scan.complexFieldOpenRuns.push(withOrphanFieldCharsPreserved(run, runElement));
      if (instrText) {
        scan.complexFieldInstr += instrText;
      }
      // Prefer any field run that actually carries formatting (the begin
      // run is sometimes an empty `<w:rPr/>` in docs that put `w:rPr` on a
      // later run). An empty formatting object counts as absent so it does
      // not block that later, genuinely-formatted run.
      const captureIsEmpty =
        !scan.complexFieldFormatting || Object.keys(scan.complexFieldFormatting).length === 0;
      if (captureIsEmpty && run.formatting && Object.keys(run.formatting).length > 0) {
        scan.complexFieldFormatting = run.formatting;
      }

      // A single physical run may contain both the instruction text and
      // structural begin/separate/end markers. Preserve only the content
      // inside the code region: fieldCode owns authored code content, while
      // the ComplexField serializer owns all structural markers.
      let inFieldCodeRegion = !hasFieldBegin && !scan.afterSeparator;
      const fieldCodeContent: Run["content"] = [];
      for (const content of run.content) {
        if (content.type === "fieldChar") {
          inFieldCodeRegion = content.charType === "begin";
          continue;
        }
        if (inFieldCodeRegion) {
          fieldCodeContent.push(content);
        }
      }
      if (fieldCodeContent.length > 0) {
        scan.complexFieldCodeRuns.push(
          fieldCodeContent.length === run.content.length
            ? run
            : { ...run, content: fieldCodeContent },
        );
      }

      if (hasFieldSeparate) {
        scan.afterSeparator = true;
      }

      if (scan.afterSeparator && !hasFieldEnd) {
        // Add to result runs (excluding the separator run itself)
        if (!hasFieldSeparate) {
          scan.complexFieldResultRuns.push(run);
        }
      }

      if (hasFieldEnd) {
        let resultRuns = scan.complexFieldResultRuns;
        // Legacy form checkboxes are rendered from `w:ffData`; they often
        // have a separator but no cached result run of their own.
        if (
          resultRuns.length === 0 &&
          scan.complexFieldFallbackDisplay !== undefined &&
          isLegacyFormCheckboxInstruction(scan.complexFieldInstr)
        ) {
          resultRuns = [
            createLegacyFormCheckboxResultRun(
              scan.complexFieldFallbackDisplay,
              scan.complexFieldFormatting,
            ),
          ];
        }
        // Self-numbering fields (LISTNUM, AUTONUM, …) often skip the
        // separator and stash their display on the end field character.
        if (resultRuns.length === 0 && !scan.afterSeparator && endOriginalValue !== undefined) {
          resultRuns = [
            {
              type: "run",
              content: [{ type: "text", text: endOriginalValue }],
            },
          ];
        }

        // Close the complex field
        const complexField: ComplexField = {
          type: "complexField",
          instruction: scan.complexFieldInstr,
          fieldType: parseFieldType(scan.complexFieldInstr),
          fieldCode: scan.complexFieldCodeRuns,
          fieldResult: resultRuns,
          ...scan.complexFieldState,
        };

        if (scan.complexFieldFormatting) {
          complexField.formatting = scan.complexFieldFormatting;
        }

        contents.push(complexField);
        if (commentReferenceId !== null) {
          contents.push({
            type: "commentReference",
            id: commentReferenceId,
          });
        }
        scan.inComplexField = false;
      }
    } else if (commentReferenceId !== null) {
      // A run whose only payload is `<w:commentReference>` parses to an
      // empty run plus the reference node. The empty run is a vestigial
      // artifact — the reference serializer re-emits its own run, so a lone
      // empty run here does not survive re-parsing and breaks round-trip
      // idempotence. Keep the run only when it also carries real content.
      if (runHoldsPayload(run)) {
        contents.push(withOrphanFieldCharsPreserved(run, runElement));
      }
      contents.push({
        type: "commentReference",
        id: commentReferenceId,
      });
    } else {
      // Regular run, not part of a field. A `separate` or `end` character
      // with no `begin` before it lands here, and it is field structure
      // with no field: the same capture keeps it.
      if (hasRunPayload({ run, runElement, rels, media })) {
        contents.push(withOrphanFieldCharsPreserved(run, runElement));
      }
    }
  },

  hyperlink: (child, { contents, styles, theme, rels, media, previews, inScopeXmlns }) => {
    contents.push(
      ...parseHyperlinkParagraphContents(
        child,
        rels,
        styles,
        theme,
        media,
        inScopeXmlns,
        previews,
      ),
    );
  },

  bookmarkStart: (child, { contents }) => {
    contents.push(parseBookmarkStart(child));
  },

  bookmarkEnd: (child, { contents }) => {
    contents.push(parseBookmarkEnd(child));
  },

  fldSimple: (child, { contents, styles, theme, rels, media, previews, inScopeXmlns }) => {
    contents.push(parseSimpleField(child, styles, theme, rels, media, inScopeXmlns, previews));
  },

  pPr: PARAGRAPH_PROPERTIES_OWNER,

  customXml: paragraphInlineWrapper("customXml"),

  proofErr: CAPTURE,
  permStart: CAPTURE,
  permEnd: CAPTURE,
  subDoc: CAPTURE,
  customXmlDelRangeEnd: CAPTURE,
  customXmlDelRangeStart: CAPTURE,
  customXmlInsRangeEnd: CAPTURE,
  customXmlInsRangeStart: CAPTURE,
  customXmlMoveFromRangeEnd: CAPTURE,
  customXmlMoveFromRangeStart: CAPTURE,
  customXmlMoveToRangeEnd: CAPTURE,
  customXmlMoveToRangeStart: CAPTURE,

  // The wrapper's own record holds these verbatim and the serializer
  // writes them back ahead of the content, so the walk that reads the
  // wrapper's children must not capture them a second time.
  smartTagPr: SMART_TAG_PROPERTIES_OWNER,
  customXmlPr: CUSTOM_XML_PROPERTIES_OWNER,

  sdt: (
    child,
    { contents, styles, theme, rels, media, previews, trackedContext, inScopeXmlns },
  ) => {
    // Structured document tag - extract properties and content
    const sdtPr = findWordprocessingChild(child, "sdtPr");
    const sdtEndPr = findWordprocessingChild(child, "sdtEndPr");
    const sdtContentEl = findWordprocessingChild(child, "sdtContent");
    if (sdtContentEl) {
      // Accumulate the `w:sdt` wrapper's own xmlns before recursing; a
      // non-canonical prefix scoped on the `w:sdt` element (not just its
      // `w:sdtContent`) must reach a captured VML `w:pict` inside the
      // content control. The recursion merges `w:sdtContent`'s own xmlns on
      // top of this.
      const sdtInScopeXmlns = mergeXmlnsDeclarations(inScopeXmlns, child);
      const sdtParsed = parseParagraphContents(
        sdtContentEl,
        styles,
        theme,
        null,
        rels,
        media,
        previews,
        trackedContext,
        sdtInScopeXmlns,
      );
      const properties = parseSdtProperties(sdtPr, sdtEndPr);
      const captured = captureSdtSiblingMarkers(child);
      if (captured.before.length > 0) {
        properties.rawSdtChildrenBeforeContent = captured.before;
      }
      if (captured.after.length > 0) {
        properties.rawSdtChildrenAfterContent = captured.after;
      }
      pushInlineSdtSegments({
        contents,
        properties,
        parsedContent: sdtParsed,
      });
    }
  },

  ins: (child, { contents, styles, theme, rels, media, previews, inScopeXmlns }) => {
    // Track change: insertion — parse content and wrap
    const insInfo = parseTrackedChangeInfo(child);
    const insContent = parseParagraphContents(
      child,
      styles,
      theme,
      null,
      rels,
      media,
      previews,
      "default",
      inScopeXmlns,
    );
    pushTrackedChangeSegments({
      contents,
      type: "insertion",
      info: insInfo,
      parsedContent: insContent,
    });
  },

  del: (child, { contents, styles, theme, rels, media, previews, inScopeXmlns }) => {
    // Track change: deletion — parse content and wrap
    const delInfo = parseTrackedChangeInfo(child);
    const delContent = parseParagraphContents(
      child,
      styles,
      theme,
      null,
      rels,
      media,
      previews,
      "deletion",
      inScopeXmlns,
    );
    pushTrackedChangeSegments({
      contents,
      type: "deletion",
      info: delInfo,
      parsedContent: delContent,
    });
  },

  moveFrom: (child, { contents, styles, theme, rels, media, previews, inScopeXmlns }) => {
    const moveFromInfo = parseTrackedChangeInfo(child);
    const moveFromContent = parseParagraphContents(
      child,
      styles,
      theme,
      null,
      rels,
      media,
      previews,
      "deletion",
      inScopeXmlns,
    );
    pushTrackedChangeSegments({
      contents,
      type: "moveFrom",
      info: moveFromInfo,
      parsedContent: moveFromContent,
    });
  },

  moveTo: (child, { contents, styles, theme, rels, media, previews, inScopeXmlns }) => {
    const moveToInfo = parseTrackedChangeInfo(child);
    const moveToContent = parseParagraphContents(
      child,
      styles,
      theme,
      null,
      rels,
      media,
      previews,
      "default",
      inScopeXmlns,
    );
    pushTrackedChangeSegments({
      contents,
      type: "moveTo",
      info: moveToInfo,
      parsedContent: moveToContent,
    });
  },

  smartTag: paragraphInlineWrapper("smartTag"),

  moveFromRangeStart: (child, { contents }) => {
    contents.push({ type: "moveFromRangeStart", ...parseMoveBookmarkMarker(child) });
  },
  moveFromRangeEnd: (child, { contents }) => {
    contents.push({ type: "moveFromRangeEnd", ...parseMarkupRangeMarker(child) });
  },
  moveToRangeStart: (child, { contents }) => {
    contents.push({ type: "moveToRangeStart", ...parseMoveBookmarkMarker(child) });
  },
  moveToRangeEnd: (child, { contents }) => {
    contents.push({ type: "moveToRangeEnd", ...parseMarkupRangeMarker(child) });
  },

  commentRangeStart: (child, { contents }) => {
    contents.push({ type: "commentRangeStart", ...parseMarkupRangeMarker(child) });
  },
  commentRangeEnd: (child, { contents }) => {
    contents.push({ type: "commentRangeEnd", ...parseMarkupRangeMarker(child) });
  },

  bdo: paragraphInlineWrapper("bdo"),
  dir: paragraphInlineWrapper("dir"),
} as const satisfies ChildHandlers<"run-level-content", ParagraphContentsWalk>;

/**
 * Parse all content within a paragraph
 *
 * Returns the parsed content and any complex fields that span multiple runs
 */
function parseParagraphContents(
  paraElement: XmlElement,
  styles: StyleMap | null,
  theme: Theme | null,
  _numbering: NumberingMap | null,
  rels: RelationshipMap | null,
  media: Map<string, MediaFile> | null,
  previews: PreviewLedger,
  trackedContext: TrackedChangeParseContext = "default",
  rootXmlns: Record<string, string> = {},
): ParagraphContent[] {
  const contents: ParagraphContent[] = [];
  // Accumulate this container's own xmlns (a paragraph or tracked-change /
  // SDT wrapper may scope non-canonical prefixes) onto the inherited set, so a
  // captured VML `w:pict` replay resolves prefixes scoped at this level too.
  const inScopeXmlns = mergeXmlnsDeclarations(rootXmlns, paraElement);

  const scan: ComplexFieldScan = {
    inComplexField: false,
    complexFieldInstr: "",
    complexFieldCodeRuns: [],
    complexFieldResultRuns: [],
    complexFieldOpenRuns: [],
    afterSeparator: false,
    complexFieldState: {},
    complexFieldFallbackDisplay: undefined,
    complexFieldFormatting: undefined,
  };

  const preserved = dispatchChildrenWithContext({
    element: paraElement,
    container: "run-level-content",
    capturePosition: () => contents.length,
    undeclared: PARAGRAPH_CONTENT_UNDECLARED,
    undeclaredNamespaces: PARAGRAPH_CONTENT_UNDECLARED_NAMESPACES,
    handlers: PARAGRAPH_CONTENT_HANDLERS,
    context: {
      styles,
      theme,
      rels,
      media,
      previews,
      trackedContext,
      inScopeXmlns,
      contents,
      scan,
    },
  });

  // The paragraph ended with a field still open: a TOC begun here and closed
  // in a later paragraph, or a `begin` nothing ever closes. Either way no
  // `ComplexField` was assembled, so the runs the state machine was holding
  // are the paragraph's content and are put back where they were read.
  if (scan.inComplexField) {
    contents.push(...scan.complexFieldOpenRuns);
  }

  // The capture is inline content in its own right, not a field on a
  // neighbour: it stands between the same two siblings in the model, in the
  // editor and in the saved part, and inside a tracked-change wrapper that
  // position is what decides whether accepting the change takes the markup
  // with it.
  return withPreservedChildren(contents, preserved, preservedInlineCapture);
}

function getCommentReferenceId(runElement: XmlElement): number | null {
  const commentReference = findChild(runElement, "w", "commentReference");
  if (!commentReference) {
    return null;
  }

  const id = Number.parseInt(getAttribute(commentReference, "w", "id") ?? "", 10);
  return Number.isFinite(id) ? id : null;
}

// ============================================================================
// MAIN PARAGRAPH PARSER
// ============================================================================

/**
 * Attribute local names `w:p` has a model field for.
 *
 * `w14:paraId` and `w14:textId` are read below and written by
 * `serializeParagraph`; `folio:reviewCarrier` is folio's own. Everything else
 * the element carried is the attribute remainder.
 */
const PARAGRAPH_ATTRIBUTES: ReadonlySet<string> = new Set(["paraId", "textId", "reviewCarrier"]);

type ParseParagraphOptions = {
  inHeaderFooter?: boolean;
  rootXmlns?: Record<string, string>;
  /** The ledger of the package this paragraph belongs to. */
  previews: PreviewLedger;
  /** Delay run merging until a source-position-dependent enrichment pass completes. */
  runConsolidation?: "immediate" | "deferred";
};

/**
 * Parse a paragraph element (w:p)
 *
 * @param node - The w:p XML element
 * @param styles - Style map for resolving style references
 * @param theme - Theme for resolving theme colors/fonts
 * @param numbering - Numbering definitions for list info
 * @param rels - Relationship map for resolving hyperlink URLs
 * @param media - Media files map for image data
 * @param options - Parsing options for context-specific body behavior; a
 *   paragraph read without them charges its previews to no package
 * @returns Parsed Paragraph object
 */
export function parseParagraph(
  node: XmlElement,
  styles: StyleMap | null,
  theme: Theme | null,
  numbering: NumberingMap | null,
  rels: RelationshipMap | null = null,
  media: Map<string, MediaFile> | null = null,
  options?: ParseParagraphOptions,
): Paragraph {
  const paragraph: Paragraph = {
    type: "paragraph",
    content: [],
  };

  // Get paragraph ID attributes (Word 2010+ uses these for collaboration).
  // OOXML types these as ST_LongHexNumber (exactly 8 hex digits); a value
  // that doesn't match is not a real Word id and downstream code (paraId
  // threading, XML serialization) must not trust it as one. Drop it rather
  // than store a malformed value — comment threading already re-derives a
  // fresh id when one is missing (see ensureThreadedCommentParaIds).
  // An id above the type's maximum is brought into range here rather than at
  // save, so the id this paragraph answers to is the id the file will carry.
  const paraId = paraIdAttribute(node);
  if (paraId && isValidHexId(paraId)) {
    paragraph.paraId = paraIdInRange(paraId);
  }

  const textId = textIdAttribute(node);
  if (textId && isValidHexId(textId)) {
    paragraph.textId = paraIdInRange(textId);
  }

  // Everything else `w:p` carried, `w:rsidR` and its family above all. The
  // three names below are the ones the reads above and `serializeParagraph`
  // own; an id folio rejected as malformed is still one of them, so it is not
  // written back from the remainder either.
  const remainder = attributeRemainder({ element: node, modelled: PARAGRAPH_ATTRIBUTES });
  if (remainder) {
    paragraph.preservedAttributes = remainder;
  }

  if (
    getAttributeByNamespaceUri(node, FOLIO_REVIEW_HISTORY_NAMESPACES, "reviewCarrier") ===
    REVIEW_CARRIERS.TERMINAL_TABLE
  ) {
    paragraph.reviewCarrier = REVIEW_CARRIERS.TERMINAL_TABLE;
  }

  if (!options?.inHeaderFooter && paragraphStartsWithRenderedPageBreak(node)) {
    paragraph.renderedPageBreakBefore = true;
  }

  // Parse paragraph properties (w:pPr)
  const pPr = findChild(node, "w", "pPr");
  if (pPr) {
    const formattingResult = parseParagraphProperties(pPr, theme);
    if (formattingResult !== undefined) {
      paragraph.formatting = formattingResult;
    }
    const propertyChangesResult = parseParagraphPropertyChanges(pPr, theme, paragraph.formatting);
    if (propertyChangesResult !== undefined) {
      paragraph.propertyChanges = propertyChangesResult;
    }

    const pPrMarkResult = parseParagraphMarkChange(pPr);
    if (pPrMarkResult !== undefined) {
      paragraph.pPrMark = pPrMarkResult;
    }

    // Check for section properties within paragraph (marks end of a section)
    const sectPr = findChild(pPr, "w", "sectPr");
    if (sectPr) {
      paragraph.sectionProperties = parseSectionProperties(sectPr);
    }
  }

  // Parse paragraph contents (runs, hyperlinks, bookmarks, fields)
  const rawContent = parseParagraphContents(
    node,
    styles,
    theme,
    numbering,
    rels,
    media,
    options?.previews ?? standalonePreviewLedger(),
    "default",
    options?.rootXmlns ?? {},
  );

  // Text-box enrichment matches model runs to source w:r elements by position.
  // Its block parsers defer this merge until that source-dependent pass ends.
  paragraph.content =
    options?.runConsolidation === "deferred" ? rawContent : consolidateParagraphContent(rawContent);

  // Compute list rendering if this is a list item.
  //
  // `w:numId` and `w:ilvl` inherit INDEPENDENTLY (ECMA-376 §17.3.1.19): a tier
  // that states only the level keeps the id it inherits, and a tier that states
  // only the id keeps the inherited level. Word writes the level-only shape
  // whenever a styled list paragraph is demoted (`<w:numPr><w:ilvl w:val="1"/>
  // </w:numPr>` with the `w:num` named by the style), so treating a direct
  // `w:numPr` as a whole replacement dropped the id and left the paragraph
  // unnumbered. The style chain itself already merges per field
  // (mergeParagraphFormatting); this is the last tier, direct over style.
  const paragraphFormatting = paragraph.formatting;
  const directNumPr = paragraphFormatting?.numPr;
  const styleNumPr =
    paragraphFormatting?.styleId && styles
      ? styles.get(paragraphFormatting.styleId)?.pPr?.numPr
      : undefined;
  // Drives indent precedence below: true when the numbering REFERENCE came from
  // the style chain, whether or not the paragraph stated its own level. A
  // paragraph that states an id of its own, the reserved cancellation
  // included, owns the reference and keeps its own indents.
  const numPrFromStyle =
    styleNumPr !== undefined && (directNumPr === undefined || directNumPr.kind === "levelOnly");
  let effectiveNumPr = directNumPr;
  if (paragraphFormatting && numPrFromStyle) {
    effectiveNumPr = mergeParagraphNumbering(styleNumPr, directNumPr);
    // Store it on the paragraph formatting so downstream code sees it, and
    // record the style tier so the serializer can drop a numPr the paragraph
    // never stated — materializing style numbering as direct <w:numPr> flips
    // Word's level-indent precedence on the saved file.
    if (effectiveNumPr !== undefined) {
      paragraphFormatting.numPr = effectiveNumPr;
    }
    paragraphFormatting.numPrFromStyle = styleNumPr;
  }

  const resolvedNumbering = resolveParagraphNumbering(effectiveNumPr);
  if (numbering) {
    if (resolvedNumbering.kind === "reference") {
      const { numId, ilvl } = resolvedNumbering;
      const level = numbering.getLevel(numId, ilvl);
      if (level) {
        const levelNumFmts: NonNullable<typeof paragraph.listRendering>["levelNumFmts"] = [];
        const levelStarts: number[] = [];
        for (let levelIndex = 0; levelIndex <= ilvl; levelIndex += 1) {
          const listLevel = numbering.getLevel(numId, levelIndex);
          levelNumFmts.push(level.isLgl || !listLevel ? "decimal" : counterFormatOf(listLevel));
          levelStarts.push(listLevel?.start ?? 1);
        }
        const listRendering: NonNullable<typeof paragraph.listRendering> = {
          level: ilvl,
          numId,
          marker: level.lvlText,
          markerTemplate: level.lvlText,
          isBullet: level.numFmt === "bullet",
          levelNumFmts,
          levelStarts,
        };
        const instance = numbering.getInstance(numId);
        const overrideForLevel = instance?.levelOverrides?.find(
          (override) => override.ilvl === ilvl,
        );
        if (instance?.abstractNumId !== undefined) {
          listRendering.abstractNumId = instance.abstractNumId;
        }
        if (overrideForLevel?.startOverride !== undefined) {
          listRendering.startOverride = overrideForLevel.startOverride;
        }
        if (level.isLgl) {
          listRendering.isLegal = true;
        }
        listRendering.numFmt = level.isLgl ? "decimal" : counterFormatOf(level);
        if (level.rPr?.hidden) {
          listRendering.markerHidden = true;
        }
        const markerFormatting = markerFormattingFromLevel(level.rPr);
        if (markerFormatting) {
          listRendering.markerFormatting = markerFormatting;
        }
        if (level.rPr?.allCaps) {
          listRendering.markerAllCaps = true;
        }
        const markerAlignment = markerAlignmentForLevel(level.lvlJc);
        if (markerAlignment !== undefined) {
          listRendering.markerAlignment = markerAlignment;
        }
        if (level.suffix) {
          listRendering.markerSuffix = level.suffix;
        }
        // Count inline LISTNUM (default-list) complex fields this paragraph
        // carries. Word advances the counter at `ilvl + 1` for each, so a
        // later sibling at that depth picks up the next letter. We also
        // fold each LISTNUM's cached display value into the marker text and
        // strip the field (plus its trailing tab, which Word swallowed into
        // the marker zone) from the inline content — that way the host
        // paragraph's marker zone reads "7.1[gap](a)" and the body text on
        // line 1 begins at the same column as the wrapped lines below.
        let implicitChildLevelAdvances = 0;
        const foldedMarkerSuffix: string[] = [];
        const filteredContent: ParagraphContent[] = [];
        let dropNextTab = false;
        // Word inserts paragraph-mark / bookmark / comment-range metadata
        // between a LISTNUM field and its trailing tab. Skip those when
        // hunting for the tab to drop, otherwise `dropNextTab` clears on
        // the metadata node and the tab survives, breaking alignment.
        const isMetadataContent = (content: ParagraphContent): boolean =>
          content.type === "bookmarkStart" ||
          content.type === "bookmarkEnd" ||
          content.type === "commentRangeStart" ||
          content.type === "commentRangeEnd" ||
          content.type === "commentReference";
        for (const content of paragraph.content) {
          if (dropNextTab && isMetadataContent(content)) {
            filteredContent.push(content);
            continue;
          }
          if (dropNextTab) {
            dropNextTab = false;
            if (
              content.type === "run" &&
              content.content.length === 1 &&
              content.content[0]?.type === "tab"
            ) {
              continue;
            }
          }
          if (content.type === "complexField") {
            const isListNum =
              content.fieldType === "LISTNUM" ||
              content.instruction.trim().toUpperCase().startsWith("LISTNUM");
            if (isListNum) {
              implicitChildLevelAdvances += 1;
              const cached = extractPlainText(content.fieldResult);
              if (cached) {
                foldedMarkerSuffix.push(cached);
              }
              dropNextTab = true;
              continue;
            }
          }
          filteredContent.push(content);
        }
        if (foldedMarkerSuffix.length > 0) {
          paragraph.content = filteredContent;
          listRendering.marker = `${listRendering.marker}\t${foldedMarkerSuffix.join(" ")}`;
          const nextLevel = numbering.getLevel(numId, ilvl + 1);
          if (
            nextLevel?.pPr?.hangingIndent === true &&
            nextLevel.pPr.indentFirstLine !== undefined
          ) {
            // `indentFirstLine` is negative for hanging indents — the
            // marker column sits at the indent's positive distance.
            const hangingTwips = -nextLevel.pPr.indentFirstLine;
            if (hangingTwips > 0) {
              listRendering.markerSecondSlotOffsetTwips = hangingTwips;
            }
          }
        }
        if (implicitChildLevelAdvances > 0) {
          listRendering.implicitChildLevelAdvances = implicitChildLevelAdvances;
        }
        paragraph.listRendering = listRendering;

        // Apply level's paragraph properties (indentation) as defaults.
        // Per OOXML spec, direct w:ind on the paragraph overrides numbering
        // level indent — only use numbering indent as fallback.
        //
        // When the numbering reference itself comes from the paragraph STYLE
        // (style pPr numPr), Word gives the style chain's own w:ind
        // precedence over the numbering level's — e.g. a "Claim" style with
        // ind left=1134 hanging=1134 referencing a level with 360/360 lays
        // out at 1134. Skip the level indents the style chain covers; the
        // toProseDoc style fallback supplies the style values. Resolution is
        // per group (left vs firstLine/hanging) so a chain that only defines
        // `left` (e.g. ListParagraph) still takes the level's hanging —
        // mirrors listAttrsFromResolvedStyle so the picker and the loader
        // resolve a style identically. Direct paragraph numPr keeps the
        // level-over-style behavior (Word's toolbar-list case).
        const chainInd = numPrFromStyle
          ? styleChainInd(paragraph.formatting?.styleId, styles)
          : { left: false, firstLine: false };
        if (level.pPr) {
          if (!paragraph.formatting) {
            paragraph.formatting = {};
          }
          const directInd = pPr ? findChild(pPr, "w", "ind") : null;
          const hasDirectLeft = hasAttributeAnySpelling(directInd, "CT_Ind @left");
          // ECMA-376 §17.3.1.12: a direct w:ind whose w:firstLine or
          // w:hanging is "0" is a no-op and must not suppress the numbering
          // level's hanging slot. Only treat non-zero direct values as
          // overrides.
          const directFirstLine = directInd
            ? parseNumericAttribute(directInd, "w", "firstLine")
            : undefined;
          const directHanging = directInd
            ? parseNumericAttribute(directInd, "w", "hanging")
            : undefined;
          const hasDirectFirstLineOrHanging =
            (directFirstLine !== undefined && directFirstLine !== 0) ||
            (directHanging !== undefined && directHanging !== 0);

          if (!hasDirectLeft && !chainInd.left && level.pPr.indentLeft !== undefined) {
            paragraph.formatting.indentLeft = level.pPr.indentLeft;
          }
          if (
            !hasDirectFirstLineOrHanging &&
            !chainInd.firstLine &&
            numberingLevelHasMarkerSlot(level)
          ) {
            if (level.pPr.indentFirstLine !== undefined) {
              paragraph.formatting.indentFirstLine = level.pPr.indentFirstLine;
            }
            if (level.pPr.hangingIndent !== undefined) {
              paragraph.formatting.hangingIndent = level.pPr.hangingIndent;
            }
          }
        }
      }
    }
  }

  if (pPr) {
    assignParagraphPropertySource(paragraph, captureParagraphPropertySource(pPr));
  }

  return paragraph;
}

/**
 * Which indent groups the basedOn chain defines: `left` (w:ind left) and
 * `firstLine` (w:ind firstLine/hanging). Walks from the given style up the
 * chain; cycles are guarded. Grouping matches listAttrsFromResolvedStyle.
 */
function styleChainInd(
  styleId: string | undefined,
  styles?: StyleMap | null,
): { left: boolean; firstLine: boolean } {
  const result = { left: false, firstLine: false };
  if (!styleId || !styles) {
    return result;
  }
  const seen = new Set<string>();
  let current: string | undefined = styleId;
  while (current && !seen.has(current)) {
    seen.add(current);
    const style = styles.get(current);
    if (!style) {
      break;
    }
    const p = style.pPr;
    if (p) {
      result.left ||= p.indentLeft !== undefined;
      result.firstLine ||= p.indentFirstLine !== undefined || p.hangingIndent !== undefined;
    }
    if (result.left && result.firstLine) {
      break;
    }
    current = style.basedOn;
  }
  return result;
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

const getRunContentText = (content: RunContent): string => {
  switch (content.type) {
    case "text":
      return content.text;
    case "tab":
      return "\t";
    case "break":
      return content.breakType === "page" ? "\f" : "\n";
    case "noBreakHyphen":
      return "\u2011";
    case "softHyphen":
      return "\u00ad";
    // Preserved markup is opaque except for the text it puts on the line:
    // `w:ruby` renders its `w:rubyBase` as the word a reader reads.
    case "preservedXml":
      return content.text;
    case "drawing":
    case "endnoteRef":
    case "fieldChar":
    case "footnoteRef":
    case "instrText":
    case "renderedPageBreak":
    case "shape":
    case "symbol":
      return "";
    default: {
      const unsupported: never = content;
      return panic(
        `Unsupported run content in plain-text extraction: ${JSON.stringify(unsupported)}`,
      );
    }
  }
};

const getRunText = (run: Run): string => {
  if (run.formatting?.hidden === true) {
    return "";
  }
  return run.content.map(getRunContentText).join("");
};

const getHyperlinkText = (hyperlink: Hyperlink): string =>
  hyperlink.children
    .map((child) => {
      switch (child.type) {
        case "run":
          return getRunText(child);
        case "bookmarkStart":
        case "bookmarkEnd":
          return "";
        // Opaque markup contributes whatever it puts on the line, which is
        // nothing except for a transparent wrapper such as `w:customXml`.
        case "preservedInline":
          return child.text;
        // A transparent wrapper puts its content on the line, unchanged.
        case "inlineWrapper":
          return child.content.map(getParagraphContentText).join("");
        default: {
          const unsupported: never = child;
          return panic(
            `Unsupported hyperlink child in plain-text extraction: ${JSON.stringify(unsupported)}`,
          );
        }
      }
    })
    .join("");

const getParagraphContentText = (content: ParagraphContent): string => {
  switch (content.type) {
    case "run":
      return getRunText(content);
    case "hyperlink":
      return getHyperlinkText(content);
    case "simpleField":
      return content.content.map(getParagraphContentText).join("");
    case "complexField":
      return content.fieldResult.map(getRunText).join("");
    case "inlineSdt":
      return content.content.map(getParagraphContentText).join("");
    case "insertion":
    case "moveTo":
    // A bidirectional wrapper changes how its text is laid out and not what
    // the text is, so plain text reads straight through it.
    case "inlineWrapper":
      return content.content.map(getParagraphContentText).join("");
    case "deletion":
    case "moveFrom":
      return "";
    case "mathEquation":
      return content.plainText ?? "";
    // Opaque markup. Its text is what it puts on the line, which a
    // transparent wrapper such as `w:customXml` has and a marker element
    // does not.
    case "preservedInline":
      return content.text;
    case "bookmarkEnd":
    case "bookmarkStart":
    case "commentRangeEnd":
    case "commentRangeStart":
    case "commentReference":
    case "moveFromRangeEnd":
    case "moveFromRangeStart":
    case "moveToRangeEnd":
    case "moveToRangeStart":
      return "";
    default: {
      const unsupported: never = content;
      return panic(
        `Unsupported paragraph content in plain-text extraction: ${JSON.stringify(unsupported)}`,
      );
    }
  }
};

/**
 * Get plain text from a paragraph
 *
 * @param paragraph - Parsed Paragraph object
 * @returns Concatenated text content
 */
export function getParagraphText(paragraph: Paragraph): string {
  return paragraph.content.map(getParagraphContentText).join("");
}

/**
 * Check if a paragraph is empty (no visible content)
 *
 * @param paragraph - Parsed Paragraph object
 * @returns true if paragraph has no visible content
 */
export function isEmptyParagraph(paragraph: Paragraph): boolean {
  return (
    getParagraphText(paragraph).trim() === "" &&
    !paragraph.content.some(
      (c) =>
        c.type === "run" && c.content.some((rc) => rc.type === "drawing" || rc.type === "shape"),
    )
  );
}

/**
 * Check if a paragraph is a list item
 *
 * @param paragraph - Parsed Paragraph object
 * @returns true if paragraph has numbering properties
 */
export function isListItem(paragraph: Paragraph): boolean {
  return paragraphNumberingReferenceId(paragraph.formatting?.numPr) !== undefined;
}

/**
 * Get the list level of a paragraph (0-8)
 *
 * @param paragraph - Parsed Paragraph object
 * @returns List level or undefined if not a list item
 */
export function getListLevel(paragraph: Paragraph): number | undefined {
  const resolved = resolveParagraphNumbering(paragraph.formatting?.numPr);
  return resolved.kind === "reference" ? resolved.ilvl : undefined;
}

/**
 * Check if paragraph has a specific style
 *
 * @param paragraph - Parsed Paragraph object
 * @param styleId - Style ID to check for
 * @returns true if paragraph has the specified style
 */
export function hasStyle(paragraph: Paragraph, styleId: string): boolean {
  return paragraph.formatting?.styleId === styleId;
}

/**
 * Check if paragraph starts with a template variable {{...}}
 *
 * @param paragraph - Parsed Paragraph object
 * @returns The variable name or null
 */
export function getTemplateVariable(paragraph: Paragraph): string | null {
  const text = getParagraphText(paragraph);
  const start = text.indexOf("{{");
  if (start === -1) {
    return null;
  }
  const end = text.indexOf("}}", start + 2);
  if (end === -1 || end === start + 2) {
    return null;
  }
  return text.slice(start + 2, end);
}
