import { isNoteReferenceId } from "../model/content";
import {
  type BlockContent,
  type BlockCustomXml,
  type BlockSdt,
  type BookmarkEnd,
  type BookmarkStart,
  type DocxPackage,
  type InlineSdt,
  type MathEquation,
  MAX_REVISION_ID,
  type Paragraph,
  type ParagraphContent,
  type PositionedBookmarkMarker,
  type PreservedInline,
  type PreservedXmlContent,
  type Run,
  type RunContent,
  type SdtProperties,
  type Table,
  type TableRow,
} from "../model/document";
import {
  MAX_TABLE_PICTURE_ID,
  type TableGeometryMove,
  type TableGeometryNativeMove,
  type TableGeometryNestedMove,
  type TableGridChange,
  type TableMoveRanges,
} from "../model/formatting";
import {
  nestedTableCopyByAnchors,
  omittedTableAnchorId,
} from "../model/nestedTableRevisionHandoffs";
import { readXml, type XmlElement } from "../xml/streamingXmlParser";
import {
  OOXML_NAMESPACE_SCOPE,
  OOXML_NS,
  resolveNamespaceUri,
  XML_NAMESPACE_URI,
} from "../xml/xmlNamespaceContext";
import { blockListEndsInParagraph, storyParagraphs } from "./blocks";
import {
  readCapturedBookmark,
  withCapturedBookmarkAliases,
} from "./capturedBookmarks";
import { structurallyEqual } from "./equality";
import {
  IDENTITY_SPACES,
  identityKeysIn,
  idKey,
  type PackageBookmark,
  type PackageBookmarkCensus,
  packageBookmarkCensus,
  packageDrawingIds,
  packageMoveRangeCensus,
  paragraphIdsIn,
  slotKey,
} from "./ids";
import {
  asParagraphContent,
  childNodes,
  type InlineNode,
  rebuildNode,
} from "./leaves";
import { stampInfo } from "./review";
import { reachesPastRemovedCell } from "./tableEditDependencies";
import { rowHasCellsOrOmissions } from "./tableGrid";
import {
  nativeMovesInParagraph,
  pairedNativeCellMoves,
} from "./tableNativeMoves";
import {
  movableTablePictureId,
  withTablePictureIdentity,
} from "./tablePictureMove";
import type { RevisionStamp } from "./types";

export type TableMoveSource = {
  paragraph: Paragraph;
  id: string;
  bookmarks: readonly PackageBookmark[];
  controls: number[];
  pictures: ReadonlySet<number>;
};
export type NestedTableMoveSource = {
  table: Table;
  paragraphs: string[];
  revisions: number[];
  ranges: TableMoveRanges[];
  bookmarks: readonly PackageBookmark[];
  insertedRows: number;
  omittedAnchorCount: 0 | 2;
  controls: number[];
  pictures: ReadonlySet<number>;
};

const movableAtom = (leaf: RunContent): boolean =>
  leaf.type === "text" ||
  leaf.type === "tab" ||
  (leaf.type === "break" &&
    (leaf.breakType === undefined || leaf.breakType === "textWrapping")) ||
  leaf.type === "symbol" ||
  leaf.type === "softHyphen" ||
  leaf.type === "noBreakHyphen" ||
  (leaf.type === "preservedXml" && movableRuby(leaf)) ||
  ((leaf.type === "footnoteRef" || leaf.type === "endnoteRef") &&
    Number.isInteger(leaf.id) &&
    isNoteReferenceId(leaf.id));

const movableRun = (run: Run, pictures: Set<number>): boolean =>
  run.content.every((leaf) => {
    if (leaf.type !== "drawing") return movableAtom(leaf);
    const id = movableTablePictureId(leaf);
    if (id === undefined || pictures.has(id)) return false;
    pictures.add(id);
    return true;
  });

const wordNamespace = (uri: string | undefined): boolean =>
  uri === OOXML_NS.w ||
  uri === "http://purl.oclc.org/ooxml/wordprocessingml/main";

const wordName = (node: XmlElement) =>
  node.type === "element" && wordNamespace(node.namespaceUri)
    ? node.name?.slice(node.name.indexOf(":") + 1)
    : undefined;

const RUBY_PROPERTIES = new Set([
  "rubyAlign",
  "hps",
  "hpsRaise",
  "hpsBaseText",
  "lid",
  "dirty",
]);
const RUBY_CHILDREN = new Map([
  ["ruby", new Set(["rubyPr", "rt", "rubyBase"])],
  ["rubyPr", RUBY_PROPERTIES],
  ["rt", new Set(["r"])],
  ["rubyBase", new Set(["r"])],
  ["r", new Set(["t"])],
]);

/** Both halves can replay a phonetic guide made of plain text runs. */
const movableRuby = ({ xml }: PreservedXmlContent): boolean => {
  const parsed = readXml(xml, {
    conformance: "ooxml",
    inheritedNamespaceScope: OOXML_NAMESPACE_SCOPE,
  });
  if (parsed.status !== "parsed" || parsed.value.elements?.length !== 1)
    return false;
  const ruby = parsed.value.elements.at(0);
  if (ruby === undefined || wordName(ruby) !== "ruby") return false;

  const attrsAreLocal = (node: XmlElement, name: string) =>
    Object.keys(node.attributes ?? {}).every((key) => {
      if (key === "xmlns" || key.startsWith("xmlns:")) return true;
      const colon = key.indexOf(":");
      if (colon < 1) return false;
      const local = key.slice(colon + 1);
      const namespace = resolveNamespaceUri(
        node.namespaceScope,
        key.slice(0, colon),
      );
      return name === "t"
        ? namespace === XML_NAMESPACE_URI && local === "space"
        : RUBY_PROPERTIES.has(name) &&
            wordNamespace(namespace) &&
            local === "val";
    });
  const visit = (node: XmlElement, parent: string): boolean => {
    if (node.type === "comment") return true;
    if (node.type === "text")
      return parent === "t" || String(node.text ?? "").trim() === "";
    const name = wordName(node);
    return (
      name !== undefined &&
      RUBY_CHILDREN.get(parent)?.has(name) === true &&
      attrsAreLocal(node, name) &&
      (node.elements?.every((child) => visit(child, name)) ?? true)
    );
  };
  return (
    attrsAreLocal(ruby, "ruby") &&
    (ruby.elements?.every((child) => visit(child, "ruby")) ?? false)
  );
};

const mathNamespace = (uri: string | undefined): boolean =>
  uri === OOXML_NS.m ||
  uri === "http://purl.oclc.org/ooxml/officeDocument/math";

// CT_RPr in a math run can carry display formatting without carrying a
// revision id or a reference to any other package record.
const MATH_RUN_FORMATTING = new Set([
  "b",
  "bCs",
  "color",
  "i",
  "iCs",
  "rFonts",
  "sz",
  "szCs",
]);

const namespaceDeclaration = (attribute: string) =>
  attribute === "xmlns" || attribute.startsWith("xmlns:");

const localMathRunProperties = (node: XmlElement): boolean =>
  Object.keys(node.attributes ?? {}).every(namespaceDeclaration) &&
  (node.elements?.every((child) => {
    const property = wordName(child);
    return (
      property !== undefined &&
      MATH_RUN_FORMATTING.has(property) &&
      (child.elements?.length ?? 0) === 0 &&
      Object.keys(child.attributes ?? {}).every((key) => {
        if (namespaceDeclaration(key)) return true;
        const colon = key.indexOf(":");
        return (
          colon > 0 &&
          key.slice(colon + 1) !== "id" &&
          wordNamespace(
            resolveNamespaceUri(child.namespaceScope, key.slice(0, colon)),
          )
        );
      })
    );
  }) ??
    true);

const localMathNode = (
  node: XmlElement,
  collectBookmark: MoveContentContext["collectBookmark"],
): boolean => {
  if (node.type === "text" || node.type === "comment") return true;
  if (node.type !== "element") return false;
  if (!mathNamespace(node.namespaceUri)) {
    if (collectBookmark === undefined || (node.elements?.length ?? 0) !== 0)
      return false;
    const bookmark = readCapturedBookmark(node);
    if (bookmark === undefined) return false;
    const names = Object.keys(node.attributes ?? {}).filter(
      (name) => name !== "xmlns" && !name.startsWith("xmlns:"),
    );
    return (
      names.length === (bookmark.marker.type === "bookmarkStart" ? 2 : 1) &&
      collectBookmark(bookmark.marker)
    );
  }
  return (
    Object.keys(node.attributes ?? {}).every((name) => {
      if (namespaceDeclaration(name)) return true;
      const colon = name.indexOf(":");
      if (colon <= 0) return false;
      const namespace = resolveNamespaceUri(
        node.namespaceScope,
        name.slice(0, colon),
      );
      return (
        mathNamespace(namespace) ||
        (namespace === XML_NAMESPACE_URI && name.slice(colon + 1) === "space")
      );
    }) &&
    (node.elements?.every((child) => {
      if (wordName(child) !== "rPr")
        return localMathNode(child, collectBookmark);
      const name = node.name ?? "";
      const owner = name.slice(name.indexOf(":") + 1);
      return (
        (owner === "r" || owner === "ctrlPr") && localMathRunProperties(child)
      );
    }) ??
      true)
  );
};

/** Admit a captured equation and its known local formatting and bookmarks. */
const movableMath = (
  { display, ommlXml }: MathEquation,
  collectBookmark: MoveContentContext["collectBookmark"],
): boolean => {
  const parsed = readXml(ommlXml, {
    conformance: "ooxml",
    inheritedNamespaceScope: OOXML_NAMESPACE_SCOPE,
  });
  if (parsed.status !== "parsed" || parsed.value.elements?.length !== 1)
    return false;
  const root = parsed.value.elements.at(0);
  const name = root?.name;
  return (
    root !== undefined &&
    name !== undefined &&
    name.slice(name.indexOf(":") + 1) ===
      (display === "inline" ? "oMath" : "oMathPara") &&
    mathNamespace(root.namespaceUri) &&
    localMathNode(root, collectBookmark)
  );
};

const readCapturedElement = (xml: string) => {
  const parsed = readXml(xml, {
    conformance: "ooxml",
    inheritedNamespaceScope: OOXML_NAMESPACE_SCOPE,
  });
  if (parsed.status !== "parsed" || parsed.value.elements?.length !== 1)
    return undefined;
  const element = parsed.value.elements.at(0);
  if (element?.type !== "element") return undefined;
  const attributes = Object.entries(element.attributes ?? {}).filter(
    ([name]) => name !== "xmlns" && !name.startsWith("xmlns:"),
  );
  const localName = element.name?.slice(element.name.indexOf(":") + 1);
  return { element, localName, attributes };
};

const emptyElement = (xml: string) => {
  const parsed = readCapturedElement(xml);
  return (parsed?.element.elements?.length ?? 0) === 0 ? parsed : undefined;
};

/** The copied cell remains in the same part and uses the same import relationship. */
const movableAltChunk = (xml: string): boolean => {
  const parsed = emptyElement(xml);
  const relationship = parsed?.attributes.at(0);
  if (
    parsed === undefined ||
    wordName(parsed.element) !== "altChunk" ||
    parsed.attributes.length !== 1 ||
    relationship === undefined ||
    typeof relationship[1] !== "string" ||
    relationship[1] === ""
  )
    return false;
  const key = relationship[0];
  const colon = key.indexOf(":");
  if (colon <= 0 || key.slice(colon + 1) !== "id") return false;
  const namespace = resolveNamespaceUri(
    parsed.element.namespaceScope,
    key.slice(0, colon),
  );
  return (
    namespace === OOXML_NS.r ||
    namespace === "http://purl.oclc.org/ooxml/officeDocument/relationships"
  );
};

/** An authored tag wrapper without a separately owned record. */
export const movableCustomXmlBlock = ({
  openingXml,
  closingXml,
}: Pick<BlockCustomXml, "openingXml" | "closingXml">): boolean => {
  const parsed = emptyElement(`${openingXml}${closingXml}`);
  return (
    parsed?.localName === "customXml" &&
    wordNamespace(parsed.element.namespaceUri) &&
    parsed.attributes.every(([attribute]) => {
      const colon = attribute.indexOf(":");
      if (colon < 1) return false;
      const name = attribute.slice(colon + 1);
      return (
        (name === "uri" || name === "element") &&
        wordNamespace(
          resolveNamespaceUri(
            parsed.element.namespaceScope,
            attribute.slice(0, colon),
          ),
        )
      );
    })
  );
};

const localCheckbox = (element: XmlElement): boolean => {
  const kinds = new Set<string>();
  return (element.elements ?? []).every((child) => {
    if (child.type === "text")
      return typeof child.text === "string" && child.text.trim() === "";
    if (
      child.type !== "element" ||
      child.namespaceUri !== OOXML_NS.w14 ||
      (child.elements?.length ?? 0) > 0
    )
      return false;
    const name = child.name?.slice(child.name.indexOf(":") + 1);
    if (
      (name !== "checked" &&
        name !== "checkedState" &&
        name !== "uncheckedState") ||
      kinds.has(name)
    )
      return false;
    const compatible = Object.keys(child.attributes ?? {}).every(
      (attribute) => {
        if (attribute === "xmlns" || attribute.startsWith("xmlns:"))
          return true;
        const colon = attribute.indexOf(":");
        if (
          colon < 1 ||
          resolveNamespaceUri(
            child.namespaceScope,
            attribute.slice(0, colon),
          ) !== OOXML_NS.w14
        )
          return false;
        const property = attribute.slice(colon + 1);
        return (
          property === "val" || (name !== "checked" && property === "font")
        );
      },
    );
    if (compatible) kinds.add(name);
    return compatible;
  });
};

const controlPropertyKind = (
  xml: string,
): "richText" | "checkbox" | "binding" | undefined => {
  const parsed = readCapturedElement(xml);
  if (parsed === undefined) return undefined;
  const { element, localName, attributes } = parsed;
  if (
    localName === "checkbox" &&
    element.namespaceUri === OOXML_NS.w14 &&
    attributes.length === 0 &&
    localCheckbox(element)
  )
    return "checkbox";
  if ((element.elements?.length ?? 0) > 0) return undefined;
  if (
    localName === "richText" &&
    wordNamespace(element.namespaceUri) &&
    attributes.length === 0
  )
    return "richText";
  if (
    localName !== "dataBinding" ||
    !(
      wordNamespace(element.namespaceUri) ||
      element.namespaceUri === OOXML_NS.w15
    )
  )
    return undefined;
  // Both review halves refer to the same package part and XPath. The
  // control IDs stay distinct while the move is pending.
  const names = new Set<string>();
  for (const [name, value] of attributes) {
    const colon = name.indexOf(":");
    if (colon < 0) return undefined;
    const namespace = resolveNamespaceUri(
      element.namespaceScope,
      name.slice(0, colon),
    );
    if (!wordNamespace(namespace) && namespace !== OOXML_NS.w15)
      return undefined;
    const local = name.slice(colon + 1);
    if (
      (local !== "storeItemID" &&
        local !== "xpath" &&
        local !== "prefixMappings") ||
      names.has(local) ||
      typeof value !== "string" ||
      (local !== "prefixMappings" && value === "")
    )
      return undefined;
    names.add(local);
  }
  return names.has("storeItemID") && names.has("xpath") ? "binding" : undefined;
};

const emptyControlEndProperties = (xml: string): boolean => {
  const parsed = emptyElement(xml);
  return (
    parsed?.localName === "sdtEndPr" &&
    wordNamespace(parsed.element.namespaceUri) &&
    parsed.attributes.length === 0
  );
};

/** Captured properties must keep their meaning in both halves of the move. */
const movableControlId = (
  properties: InlineSdt["properties"],
): number | undefined => {
  const { id, preserved, rawEndPropertiesXml } = properties;
  const checkbox = properties.sdtType === "checkbox";
  if (
    (properties.sdtType !== "richText" && !checkbox) ||
    id === undefined ||
    !Number.isInteger(id) ||
    id < -MAX_REVISION_ID - 1 ||
    id > MAX_REVISION_ID ||
    properties.rawSdtChildrenBeforeContent !== undefined ||
    properties.rawSdtChildrenAfterContent !== undefined ||
    properties.endProperties?.runProperties !== undefined ||
    (rawEndPropertiesXml !== undefined &&
      !emptyControlEndProperties(rawEndPropertiesXml))
  )
    return undefined;
  const kinds = new Set<string>();
  for (const { xml } of preserved?.children ?? []) {
    const kind = controlPropertyKind(xml);
    if (
      kind === undefined ||
      kinds.has(kind) ||
      (kind === "checkbox" && !checkbox) ||
      (kind === "richText" && checkbox)
    )
      return undefined;
    kinds.add(kind);
  }
  return id;
};

const movableUnlockedControlId = (properties: SdtProperties) =>
  properties.lock === undefined || properties.lock === "unlocked"
    ? movableControlId(properties)
    : undefined;

export const movableBlockControl = ({ properties }: BlockSdt): boolean =>
  movableUnlockedControlId(properties) !== undefined;

type ProofingKind = "spelling" | "grammar";

/** Proofing ranges pair by kind and order. Both endpoints must move together. */
const recordProofingMarker = (
  { xml, text, foldedListNumber }: PreservedInline,
  open: Set<ProofingKind>,
): boolean => {
  if (text !== "" || foldedListNumber !== undefined) return false;
  const parsed = emptyElement(xml);
  const attribute = parsed?.attributes.at(0);
  if (
    parsed?.localName !== "proofErr" ||
    !wordNamespace(parsed.element.namespaceUri) ||
    parsed.attributes.length !== 1 ||
    attribute === undefined
  )
    return false;
  const [name, value] = attribute;
  const colon = name.indexOf(":");
  if (
    colon <= 0 ||
    name.slice(colon + 1) !== "type" ||
    !wordNamespace(
      resolveNamespaceUri(parsed.element.namespaceScope, name.slice(0, colon)),
    )
  )
    return false;
  let kind: ProofingKind;
  switch (value) {
    case "spellEnd":
      return open.delete("spelling");
    case "gramEnd":
      return open.delete("grammar");
    case "spellStart":
      kind = "spelling";
      break;
    case "gramStart":
      kind = "grammar";
      break;
    default:
      return false;
  }
  if (open.has(kind)) return false;
  open.add(kind);
  return true;
};

type MoveContentContext = {
  withinLink: boolean;
  controls: number[];
  proofing: Set<ProofingKind>;
  pictures: Set<number>;
  priorRevisions?: number[];
  collectBookmark?: (marker: BookmarkStart | BookmarkEnd) => boolean;
};

const nestedMoveContext = (context: MoveContentContext): MoveContentContext => {
  if (context.priorRevisions === undefined) return context;
  const { priorRevisions: _revisions, ...nested } = context;
  return nested;
};

// Both sides of the move carry the inline content. Admit only records whose
// identity and ownership stay local to the paragraph.
const movableInlineContent = (
  content: readonly ParagraphContent[],
  context: MoveContentContext,
): boolean =>
  content.length > 0 &&
  content.every((item) => {
    switch (item.type) {
      case "run":
        if (item.propertyChanges !== undefined) {
          if (context.priorRevisions === undefined) return false;
          for (const { info } of item.propertyChanges)
            context.priorRevisions.push(info.id);
        }
        return movableRun(item, context.pictures);
      case "bookmarkStart":
      case "bookmarkEnd":
        return context.collectBookmark?.(item) ?? false;
      case "insertion":
      case "deletion": {
        if (
          context.priorRevisions === undefined ||
          item.resolutionJoins !== undefined
        )
          return false;
        context.priorRevisions.push(item.info.id);
        // The nested table's native row owns the copy. Keep these older
        // top-level run records independently reviewable in both copies.
        return movableInlineContent(item.content, nestedMoveContext(context));
      }
      case "mathEquation":
        return movableMath(item, context.collectBookmark);
      case "preservedInline":
        return recordProofingMarker(item, context.proofing);
      case "hyperlink":
        // A bidi wrapper can contain a link, but cannot hide a nested one.
        return (
          !context.withinLink &&
          movableInlineContent(item.children, {
            ...context,
            withinLink: true,
          })
        );
      case "inlineWrapper":
        return (
          (item.kind === "bidi" || item.propertiesXml === undefined) &&
          movableInlineContent(item.content, context)
        );
      case "inlineSdt": {
        const id = movableControlId(item.properties);
        if (id === undefined) return false;
        context.controls.push(id);
        return movableInlineContent(item.content, context);
      }
      default:
        return false;
    }
  });

/** The run move can take the paragraph whole without duplicating an anchor. */
export const tableMoveSource = (
  paragraph: Paragraph,
  nested?: Pick<MoveContentContext, "priorRevisions" | "collectBookmark"> & {
    proofing?: Set<ProofingKind>;
  },
): TableMoveSource | undefined => {
  const controls: number[] = [];
  const proofing = nested?.proofing ?? new Set<ProofingKind>();
  let hasLocalBookmarks = false;
  const collectBookmark =
    nested?.collectBookmark ??
    ((marker: BookmarkStart | BookmarkEnd) => {
      hasLocalBookmarks = true;
      return movableBookmark(marker);
    });
  const pictures = new Set<number>();
  if (
    paragraph.paraId === undefined ||
    paragraph.pPrMark !== undefined ||
    (paragraph.propertyChanges !== undefined &&
      (nested?.priorRevisions === undefined ||
        paragraph.propertyChanges.some(
          ({ resolvesWith }) => resolvesWith !== undefined,
        ))) ||
    paragraph.sectionProperties !== undefined ||
    paragraph.reviewCarrier !== undefined ||
    (paragraph.content.length !== 0 &&
      !movableInlineContent(paragraph.content, {
        withinLink: false,
        controls,
        pictures,
        ...nested,
        proofing,
        collectBookmark,
      })) ||
    (nested?.proofing === undefined && proofing.size !== 0)
  )
    return undefined;
  const bookmarks = hasLocalBookmarks ? bookmarksIn(paragraph) : undefined;
  if (bookmarks !== undefined && !completeBookmarks(bookmarks))
    return undefined;
  for (const { info } of paragraph.propertyChanges ?? [])
    nested?.priorRevisions?.push(info.id);
  return {
    paragraph,
    id: paragraph.paraId,
    bookmarks: bookmarks?.pairs ?? [],
    controls,
    pictures,
  };
};

type MoveResourcePairs = readonly (readonly [number, number])[];
type MoveResourceIdsOptions = {
  values: ReadonlySet<number>;
  pairs: MoveResourcePairs | undefined;
  side: 0 | 1;
};
const hasMoveResourceIds = ({
  values,
  pairs,
  side,
}: MoveResourceIdsOptions): boolean =>
  values.size === (pairs?.length ?? 0) &&
  (pairs?.every((pair) => values.has(pair[side])) ?? true);

type MoveBookmarkPairs = NonNullable<TableGeometryNestedMove["bookmarks"]>;

const bookmarksIn = (block: BlockContent) =>
  packageBookmarkCensus({ document: { content: [block] } });

const completeBookmarks = ({ ids, pairs }: PackageBookmarkCensus): boolean =>
  ids.size === pairs.length &&
  new Set(pairs.map(({ name }) => name.toLowerCase())).size === pairs.length;

const movableBookmark = (marker: BookmarkStart | BookmarkEnd): boolean => {
  const bookmarkId = marker.id;
  return (
    Number.isInteger(bookmarkId) &&
    bookmarkId >= 0 &&
    bookmarkId <= MAX_REVISION_ID &&
    marker.copyOriginAncestors === undefined &&
    (marker.type === "bookmarkEnd" || marker.name.length > 0)
  );
};

type NativeHalves = NonNullable<
  ReturnType<typeof nativeMovesInParagraph>
>["halves"];

/** A structural copy may carry only records whose package identities stay unique. */
export const nestedTableMoveSource = (
  table: Table,
): NestedTableMoveSource | undefined => {
  let hasBookmarks = false;
  const proofing = new Set<ProofingKind>();
  const collectBookmark = (marker: BookmarkStart | BookmarkEnd) => {
    hasBookmarks = true;
    return movableBookmark(marker);
  };
  const paragraphs: string[] = [];
  const revisions: number[] = [];
  const ranges: TableMoveRanges[] = [];
  const nativeHalves: NativeHalves = [];
  const earlierMoves = new Map<string, TableGeometryMove>();
  const gridMembers: number[] = [];
  const controls = new Set<number>();
  const pictures = new Set<number>();
  const collectControl = (properties: SdtProperties): boolean => {
    const id = movableUnlockedControlId(properties);
    if (id === undefined) return false;
    controls.add(id);
    return true;
  };
  const collectCarriers = ({
    contentControls,
    carrierStack,
  }: Pick<TableRow, "contentControls" | "carrierStack">): boolean =>
    (contentControls?.every(collectControl) ?? true) &&
    (carrierStack?.every((carrier) =>
      carrier.type === "customXml"
        ? movableCustomXmlBlock(carrier.wrapper)
        : collectControl(carrier.properties),
    ) ??
      true);
  const collectPositionedBookmarks = (
    bookmarks: PositionedBookmarkMarker[] | undefined,
  ) =>
    bookmarks?.every(
      (bookmark) =>
        movableBookmark(bookmark.marker) && collectCarriers(bookmark),
    ) ?? true;
  const collectProperties = (
    changes: readonly { info: { id: number } }[] | undefined,
  ) => {
    for (const { info } of changes ?? []) revisions.push(info.id);
  };
  const collectBlocks = (blocks: readonly BlockContent[]): boolean =>
    blocks.every((block, index) => collectBlock(block, blocks.at(index + 1)));
  const collect = (candidate: Table): boolean => {
    const { rows, bookmarks, ...tableProperties } = candidate;
    const grid = tableProperties.formatting?.gridChange;
    const geometry = grid?.geometryChange;
    hasBookmarks ||= (bookmarks?.length ?? 0) > 0;
    if (
      !collectPositionedBookmarks(bookmarks) ||
      tableProperties.carrierStack !== undefined ||
      reachesPastRemovedCell(tableProperties) ||
      (geometry?.type === "mergeCells" &&
        !tableMoveOwnershipMatches(candidate, geometry)) ||
      rows.length === 0
    )
      return false;
    if (grid !== undefined) {
      revisions.push(grid.id);
      for (const id of geometry?.revisionIds ?? []) gridMembers.push(id);
    }
    if (geometry?.type === "mergeCells") {
      for (const move of geometry.moves ?? []) {
        earlierMoves.set(idKey(move.source), move);
        earlierMoves.set(idKey(move.destination), move);
      }
    }
    collectProperties(tableProperties.propertyChanges);
    return rows.every((row) => {
      const {
        cells,
        bookmarks: rowBookmarks,
        contentControls: _rowControls,
        carrierStack: _rowStack,
        ...rowProperties
      } = row;
      hasBookmarks ||= (rowBookmarks?.length ?? 0) > 0;
      if (
        !collectPositionedBookmarks(rowBookmarks) ||
        !collectCarriers(row) ||
        reachesPastRemovedCell(rowProperties) ||
        !rowHasCellsOrOmissions(row)
      )
        return false;
      if (row.structuralChange !== undefined)
        revisions.push(row.structuralChange.info.id);
      collectProperties(row.propertyChanges);
      collectProperties(row.tablePropertyExceptionChanges);
      return cells.every((cell) => {
        const {
          content,
          contentControls: _cellControls,
          carrierStack: _cellStack,
          ...cellProperties
        } = cell;
        if (
          !collectCarriers(cell) ||
          reachesPastRemovedCell(cellProperties) ||
          !blockListEndsInParagraph(content)
        )
          return false;
        const change = cell.structuralChange;
        if (change !== undefined) {
          if (
            change.type === "tableRowInsertion" ||
            change.type === "tableRowDeletion"
          )
            return false;
          revisions.push(change.info.id);
        }
        collectProperties(cell.propertyChanges);
        for (const { previousStructuralChange } of cell.propertyChanges ?? []) {
          if (previousStructuralChange === undefined) continue;
          if (
            previousStructuralChange.type === "tableRowInsertion" ||
            previousStructuralChange.type === "tableRowDeletion"
          )
            return false;
          revisions.push(previousStructuralChange.info.id);
        }
        return collectBlocks(content);
      });
    });
  };
  const collectBlock = (
    block: BlockContent,
    following: BlockContent | undefined,
  ): boolean => {
    if (block.type === "bookmarkStart" || block.type === "bookmarkEnd")
      return collectBookmark(block);
    if (block.type === "table") return collect(block);
    if (block.type === "preservedBlock") return movableAltChunk(block.xml);
    if (block.type === "blockCustomXml")
      return movableCustomXmlBlock(block) && collectBlocks(block.content);
    if (block.type === "blockSdt") {
      return collectControl(block.properties) && collectBlocks(block.content);
    }
    if (block.type !== "paragraph") return false;
    const move =
      block.paraId === undefined
        ? undefined
        : earlierMoves.get(idKey(block.paraId));
    let plain = block;
    if (move !== undefined) {
      const { pPrMark, ...paragraph } = block;
      if (pPrMark !== undefined) {
        if (idKey(block.paraId ?? "") !== idKey(move.destination)) return false;
        revisions.push(pPrMark.info.id);
      }
      plain = paragraph;
      if (move.kind !== "empty") {
        const [start, tracked, end] = block.content;
        if (
          (start?.type !== "moveFromRangeStart" &&
            start?.type !== "moveToRangeStart") ||
          (tracked?.type !== "moveFrom" && tracked?.type !== "moveTo") ||
          (end?.type !== "moveFromRangeEnd" &&
            end?.type !== "moveToRangeEnd") ||
          start.copyOriginAncestors !== undefined ||
          end.copyOriginAncestors !== undefined ||
          tracked.resolutionJoins !== undefined
        )
          return false;
        revisions.push(tracked.info.id);
        if (start.type === "moveFromRangeStart")
          ranges.push({
            fromRange: move.fromRange,
            toRange: move.toRange,
            name: start.name,
          });
        plain = { ...paragraph, content: tracked.content };
      }
    } else if (block.pPrMark !== undefined) {
      const { pPrMark, ...paragraph } = block;
      if (
        // These ordinary native marks join into the next paragraph in the
        // same copied block list. Folio cut or handoff metadata needs its own
        // remapping, as do moves that do not belong to this table's history.
        following?.type !== "paragraph" ||
        following.paraId === undefined ||
        (pPrMark.kind !== "ins" && pPrMark.kind !== "del") ||
        pPrMark.deletion !== undefined ||
        pPrMark.resolutionJoin !== undefined ||
        pPrMark.splitDate !== undefined ||
        pPrMark.survivor !== undefined ||
        pPrMark.takenOverMark !== undefined
      )
        return false;
      revisions.push(pPrMark.info.id);
      plain = paragraph;
    } else {
      const native = nativeMovesInParagraph(block.content);
      if (native === undefined) return false;
      for (const half of native.halves) {
        revisions.push(half.revision);
        nativeHalves.push(half);
      }
      if (native.halves.length > 0)
        plain = { ...block, content: native.content };
    }
    const source = tableMoveSource(plain, {
      ...(move === undefined ? { priorRevisions: revisions, proofing } : {}),
      collectBookmark,
    });
    if (source === undefined) return false;
    if (move !== undefined && move.kind !== "empty") {
      const side = idKey(source.id) === idKey(move.source) ? 0 : 1;
      if (
        !hasMoveResourceIds({
          values: new Set(source.controls),
          pairs: move.controls,
          side,
        }) ||
        !hasMoveResourceIds({
          values: source.pictures,
          pairs: move.pictures,
          side,
        })
      )
        return false;
    }
    paragraphs.push(source.id);
    for (const id of source.controls) controls.add(id);
    for (const id of source.pictures) {
      if (pictures.has(id)) return false;
      pictures.add(id);
    }
    return true;
  };
  const insertedRows = table.rows.filter(
    ({ structuralChange }) => structuralChange === undefined,
  ).length;
  if (!collect(table) || proofing.size !== 0) return undefined;
  const nativeMoves = pairedNativeCellMoves(nativeHalves);
  if (nativeMoves === undefined) return undefined;
  for (const { name, fromRange, toRange } of nativeMoves)
    ranges.push({ name, fromRange, toRange });
  const bookmarks = hasBookmarks ? bookmarksIn(table) : undefined;
  if (bookmarks !== undefined && !completeBookmarks(bookmarks))
    return undefined;
  const revisionIds = new Set(revisions);
  if (gridMembers.some((id) => !revisionIds.has(id))) return undefined;
  const allowed = new Set([
    ...revisions.map((id) => slotKey({ space: IDENTITY_SPACES.REVISION, id })),
    ...[...controls].map((id) =>
      slotKey({ space: IDENTITY_SPACES.CONTROL, id }),
    ),
  ]);
  const actual = identityKeysIn(table);
  return revisions.length + controls.size === allowed.size &&
    actual.length === allowed.size &&
    actual.every((key) => allowed.has(key))
    ? {
        table,
        paragraphs,
        revisions,
        ranges,
        bookmarks: bookmarks?.pairs ?? [],
        controls: [...controls],
        pictures,
        insertedRows,
        omittedAnchorCount:
          paragraphs.length === 0 && insertedRows === 0 ? 2 : 0,
      }
    : undefined;
};

type NestedMoveIdentities = {
  paragraphs: ReadonlyMap<string, string>;
  revisions: ReadonlyMap<number, number>;
  controls: ReadonlyMap<number, number>;
  pictures: ReadonlyMap<number, number>;
  ranges: ReadonlyMap<number, { id: number; name: string }>;
  bookmarks: ReadonlyMap<number, PackageBookmark>;
  bookmarkNames: ReadonlyMap<string, string>;
};

const bookmarkAliases = (pairs: MoveBookmarkPairs, reverse: boolean) => {
  const bookmarks = new Map<number, PackageBookmark>();
  const bookmarkNames = new Map<string, string>();
  for (const pair of pairs) {
    const [before, after] = reverse ? [pair[1], pair[0]] : pair;
    bookmarks.set(before.id, after);
    bookmarkNames.set(before.name, after.name);
  }
  return { bookmarks, bookmarkNames };
};

const bookmarkOwnershipMatches = (
  source: BlockContent,
  target: BlockContent,
  aliases: MoveBookmarkPairs | undefined,
): boolean => {
  if (aliases === undefined) return true;
  const originals = bookmarksIn(source);
  const copies = bookmarksIn(target);
  // Resolving an earlier move can remove the corresponding pair from both
  // nested copies while their containing move is still pending.
  const active = aliases.filter(
    ([from, to]) => originals.ids.has(from.id) || copies.ids.has(to.id),
  );
  return (
    originals.ids.size === active.length &&
    copies.ids.size === active.length &&
    new Set(active.map(([from]) => from.id)).size === active.length &&
    new Set(active.map(([, to]) => to.id)).size === active.length &&
    active.every(
      ([from, to]) =>
        from.id !== to.id &&
        from.name.toLowerCase() !== to.name.toLowerCase() &&
        originals.pairs.some(
          ({ id, name }) => id === from.id && name === from.name,
        ) &&
        copies.pairs.some(({ id, name }) => id === to.id && name === to.name),
    )
  );
};

type NestedMoveRanges = NonNullable<TableGeometryNestedMove["moveRanges"]>;
const nestedMoveRangeMap = (
  pairs: NestedMoveRanges,
  reverse: boolean,
): NestedMoveIdentities["ranges"] => {
  const map = new Map<number, { id: number; name: string }>();
  for (const [original, temporary] of pairs) {
    const before = reverse ? temporary : original;
    const after = reverse ? original : temporary;
    map.set(before.fromRange, { id: after.fromRange, name: after.name });
    map.set(before.toRange, { id: after.toRange, name: after.name });
  }
  return map;
};

const withNestedMoveIdentities = (
  table: Table,
  ids: NestedMoveIdentities,
): Table => {
  const restampControl = (properties: SdtProperties): SdtProperties => {
    const id =
      properties.id === undefined ? undefined : ids.controls.get(properties.id);
    return id === undefined ? properties : { ...properties, id };
  };
  const carrierIdentities = <
    T extends Pick<TableRow, "contentControls" | "carrierStack">,
  >(
    item: T,
  ): T => ({
    ...item,
    ...(item.contentControls === undefined
      ? {}
      : {
          contentControls: item.contentControls.map(restampControl),
        }),
    ...(item.carrierStack === undefined
      ? {}
      : {
          carrierStack: item.carrierStack.map((carrier) =>
            carrier.type === "sdt"
              ? { ...carrier, properties: restampControl(carrier.properties) }
              : carrier,
          ),
        }),
  });
  const rewriteBookmark = (
    bookmark: PositionedBookmarkMarker,
  ): PositionedBookmarkMarker => {
    const marker = withBookmarkIdentity(bookmark.marker, ids.bookmarks);
    return { ...carrierIdentities(bookmark), marker };
  };
  const restamp = <Change extends { info: { id: number } }>(
    record: Change,
  ): Change => {
    const id = ids.revisions.get(record.info.id);
    return id === undefined
      ? record
      : { ...record, info: { ...record.info, id } };
  };
  const rewriteMoveContent = (content: InlineNode): InlineNode => {
    switch (content.type) {
      case "insertion":
      case "deletion":
      case "moveFrom":
      case "moveTo":
        return restamp(content);
      case "run":
        return content.propertyChanges === undefined
          ? content
          : {
              ...content,
              propertyChanges: content.propertyChanges.map(restamp),
            };
      case "moveFromRangeStart":
      case "moveToRangeStart": {
        const range = ids.ranges.get(content.id);
        return range === undefined ? content : { ...content, ...range };
      }
      case "moveFromRangeEnd":
      case "moveToRangeEnd": {
        const range = ids.ranges.get(content.id);
        return range === undefined ? content : { ...content, id: range.id };
      }
      default:
        return content;
    }
  };
  const rewrite = (block: BlockContent): BlockContent => {
    if (block.type === "table") return withNestedMoveIdentities(block, ids);
    if (block.type === "bookmarkStart" || block.type === "bookmarkEnd")
      return withBookmarkIdentity(block, ids.bookmarks);
    if (block.type === "blockCustomXml")
      return { ...block, content: block.content.map(rewrite) };
    if (block.type === "blockSdt") {
      return {
        ...block,
        properties: restampControl(block.properties),
        content: block.content.map(rewrite),
      };
    }
    if (block.type !== "paragraph" || block.paraId === undefined) return block;
    const paraId = ids.paragraphs.get(idKey(block.paraId));
    return paraId === undefined
      ? block
      : {
          ...block,
          paraId,
          ...(block.pPrMark === undefined
            ? {}
            : { pPrMark: restamp(block.pPrMark) }),
          ...(block.propertyChanges === undefined
            ? {}
            : { propertyChanges: block.propertyChanges.map(restamp) }),
          content: withMoveIdentities({
            content: block.content,
            ...(ids.ranges.size === 0 && ids.revisions.size === 0
              ? {}
              : { reviewIdentity: rewriteMoveContent }),
            controls: ids.controls,
            pictures: ids.pictures,
            bookmarkNames: ids.bookmarkNames,
            bookmarkIds: ids.bookmarks,
          }),
        };
  };
  const grid = table.formatting?.gridChange;
  const gridId = grid === undefined ? undefined : ids.revisions.get(grid.id);
  const revisionId = (id: number) => ids.revisions.get(id) ?? id;
  const paragraphId = (id: string) => ids.paragraphs.get(idKey(id)) ?? id;
  const rewriteResources = (
    pairs: MoveResourcePairs,
    aliases: ReadonlyMap<number, number>,
  ) =>
    pairs.map(
      ([from, to]) =>
        [aliases.get(from) ?? from, aliases.get(to) ?? to] as const,
    );
  const rewriteBookmarkPairs = (pairs: MoveBookmarkPairs): MoveBookmarkPairs =>
    pairs.map(([from, to]) => [
      ids.bookmarks.get(from.id) ?? from,
      ids.bookmarks.get(to.id) ?? to,
    ]);
  const rewriteMove = (move: TableGeometryMove): TableGeometryMove => {
    const addresses = {
      source: paragraphId(move.source),
      destination: paragraphId(move.destination),
      destinationMark: revisionId(move.destinationMark),
    };
    if (move.kind === "empty") return { ...addresses, kind: "empty" };
    return {
      ...move,
      ...addresses,
      fromRevision: revisionId(move.fromRevision),
      toRevision: revisionId(move.toRevision),
      fromRange: ids.ranges.get(move.fromRange)?.id ?? move.fromRange,
      toRange: ids.ranges.get(move.toRange)?.id ?? move.toRange,
      ...(move.controls === undefined
        ? {}
        : { controls: rewriteResources(move.controls, ids.controls) }),
      ...(move.pictures === undefined
        ? {}
        : { pictures: rewriteResources(move.pictures, ids.pictures) }),
      ...(move.bookmarks === undefined
        ? {}
        : { bookmarks: rewriteBookmarkPairs(move.bookmarks) }),
    };
  };
  const rewriteRanges = (ranges: TableMoveRanges): TableMoveRanges => {
    const from = ids.ranges.get(ranges.fromRange);
    const to = ids.ranges.get(ranges.toRange);
    return {
      fromRange: from?.id ?? ranges.fromRange,
      toRange: to?.id ?? ranges.toRange,
      name: from?.name ?? to?.name ?? ranges.name,
    };
  };
  const rewriteNested = (
    move: TableGeometryNestedMove,
  ): TableGeometryNestedMove => ({
    rows: move.rows.map(revisionId),
    paragraphs: move.paragraphs.map(([from, to]) => [
      paragraphId(from),
      paragraphId(to),
    ]),
    ...(move.omittedTable === undefined
      ? {}
      : {
          omittedTable: [
            revisionId(move.omittedTable[0]),
            revisionId(move.omittedTable[1]),
          ] as const,
        }),
    ...(move.revisions === undefined
      ? {}
      : { revisions: rewriteResources(move.revisions, ids.revisions) }),
    ...(move.controls === undefined
      ? {}
      : { controls: rewriteResources(move.controls, ids.controls) }),
    ...(move.pictures === undefined
      ? {}
      : { pictures: rewriteResources(move.pictures, ids.pictures) }),
    ...(move.moveRanges === undefined
      ? {}
      : {
          moveRanges: move.moveRanges.map(
            ([from, to]) => [rewriteRanges(from), rewriteRanges(to)] as const,
          ),
        }),
    ...(move.bookmarks === undefined
      ? {}
      : { bookmarks: rewriteBookmarkPairs(move.bookmarks) }),
  });
  const geometry = grid?.geometryChange;
  const rewrittenGeometry =
    geometry === undefined
      ? undefined
      : {
          ...geometry,
          revisionIds: geometry.revisionIds.map(revisionId),
          ...(geometry.type !== "mergeCells"
            ? {}
            : {
                ...(geometry.moves === undefined
                  ? {}
                  : { moves: geometry.moves.map(rewriteMove) }),
                ...(geometry.nestedMoves === undefined
                  ? {}
                  : { nestedMoves: geometry.nestedMoves.map(rewriteNested) }),
                ...(geometry.emptyOwner === undefined
                  ? {}
                  : { emptyOwner: paragraphId(geometry.emptyOwner) }),
              }),
        };
  return {
    ...table,
    ...(grid === undefined || gridId === undefined
      ? {}
      : {
          formatting: {
            ...table.formatting,
            gridChange: {
              ...grid,
              id: gridId,
              ...(rewrittenGeometry === undefined
                ? {}
                : { geometryChange: rewrittenGeometry }),
            },
          },
        }),
    ...(table.propertyChanges === undefined
      ? {}
      : { propertyChanges: table.propertyChanges.map(restamp) }),
    ...(table.bookmarks === undefined
      ? {}
      : { bookmarks: table.bookmarks.map(rewriteBookmark) }),
    rows: table.rows.map((row) => ({
      ...carrierIdentities(row),
      ...(row.bookmarks === undefined
        ? {}
        : { bookmarks: row.bookmarks.map(rewriteBookmark) }),
      ...(row.structuralChange === undefined
        ? {}
        : { structuralChange: restamp(row.structuralChange) }),
      ...(row.propertyChanges === undefined
        ? {}
        : { propertyChanges: row.propertyChanges.map(restamp) }),
      ...(row.tablePropertyExceptionChanges === undefined
        ? {}
        : {
            tablePropertyExceptionChanges:
              row.tablePropertyExceptionChanges.map(restamp),
          }),
      cells: row.cells.map((cell) => ({
        ...carrierIdentities(cell),
        ...(cell.propertyChanges === undefined
          ? {}
          : {
              propertyChanges: cell.propertyChanges.map((change) => {
                const updated = restamp(change);
                return change.previousStructuralChange === undefined
                  ? updated
                  : {
                      ...updated,
                      previousStructuralChange: restamp(
                        change.previousStructuralChange,
                      ),
                    };
              }),
            }),
        ...(cell.structuralChange === undefined
          ? {}
          : { structuralChange: restamp(cell.structuralChange) }),
        content: cell.content.map(rewrite),
      })),
    })),
  };
};

type MarkNestedTableMoveOptions = {
  source: NestedTableMoveSource;
  destinations: readonly string[];
  rowStamps: readonly RevisionStamp[];
  recordStamps: readonly RevisionStamp[];
  anchorStamps: readonly RevisionStamp[];
  moveRanges: NestedMoveRanges;
  bookmarks: MoveBookmarkPairs;
  controls: readonly (readonly [number, number])[];
  pictures: readonly (readonly [number, number])[];
};
type MarkedNestedTableMove = {
  original: Table;
  destination: Table;
  record: TableGeometryNestedMove;
};
const addOmittedTableAnchor = (table: Table, stamp: RevisionStamp): Table => {
  const { gridChange: _grid, ...previousFormatting } = table.formatting ?? {};
  return {
    ...table,
    propertyChanges: [
      {
        type: "tablePropertyChange",
        info: stampInfo(stamp),
        ...(Object.keys(previousFormatting).length === 0
          ? {}
          : { previousFormatting }),
      },
      ...(table.propertyChanges ?? []),
    ],
  };
};
export const markNestedTableMove = ({
  source,
  destinations,
  rowStamps,
  recordStamps,
  anchorStamps,
  moveRanges,
  bookmarks,
  controls,
  pictures,
}: MarkNestedTableMoveOptions): MarkedNestedTableMove | undefined => {
  if (
    destinations.length !== source.paragraphs.length ||
    rowStamps.length !== source.insertedRows ||
    recordStamps.length !== source.revisions.length ||
    anchorStamps.length !== source.omittedAnchorCount
  )
    return undefined;
  const paragraphs: [string, string][] = [];
  for (const [index, original] of source.paragraphs.entries()) {
    const destination = destinations[index];
    if (destination === undefined) return undefined;
    paragraphs.push([original, destination]);
  }
  const revisions: [number, number][] = [];
  for (const [index, original] of source.revisions.entries()) {
    const temporary = recordStamps[index];
    if (temporary === undefined) return undefined;
    revisions.push([original, temporary.id]);
  }
  const table = withNestedMoveIdentities(source.table, {
    paragraphs: new Map(
      paragraphs.map(([original, destination]) => [
        idKey(original),
        destination,
      ]),
    ),
    revisions: new Map(revisions),
    controls: new Map(controls),
    pictures: new Map(pictures),
    ranges: nestedMoveRangeMap(moveRanges, false),
    ...bookmarkAliases(bookmarks, false),
  });
  let nextStamp = 0;
  const rows = table.rows.map((row) => {
    if (row.structuralChange !== undefined) return row;
    const stamp = rowStamps[nextStamp++];
    return stamp === undefined
      ? row
      : {
          ...row,
          structuralChange: {
            type: "tableRowInsertion" as const,
            info: stampInfo(stamp),
          },
        };
  });
  const destination = { ...table, rows };
  const [from, to] = anchorStamps;
  return {
    original:
      from === undefined
        ? source.table
        : addOmittedTableAnchor(source.table, from),
    destination:
      to === undefined ? destination : addOmittedTableAnchor(destination, to),
    record: {
      rows: rowStamps.map(({ id }) => id),
      paragraphs,
      ...(from === undefined || to === undefined
        ? {}
        : { omittedTable: [from.id, to.id] as const }),
      ...(revisions.length === 0 ? {} : { revisions }),
      ...(moveRanges.length === 0 ? {} : { moveRanges }),
      ...(bookmarks.length === 0 ? {} : { bookmarks }),
      ...(controls.length === 0 ? {} : { controls }),
      ...(pictures.length === 0 ? {} : { pictures }),
    } satisfies TableGeometryNestedMove,
  };
};

type MarkTableMoveOptions = {
  source: TableMoveSource;
  destination: string;
  from: RevisionStamp;
  to: RevisionStamp;
  mark: RevisionStamp;
  ranges: TableMoveRanges;
  bookmarks: MoveBookmarkPairs;
  controls: readonly (readonly [number, number])[];
  pictures: readonly (readonly [number, number])[];
};

type MoveIdentityOptions = {
  content: ParagraphContent[];
  controls: ReadonlyMap<number, number>;
  pictures: ReadonlyMap<number, number>;
  bookmarkNames?: ReadonlyMap<string, string>;
  bookmarkIds?: ReadonlyMap<number, PackageBookmark>;
  reviewIdentity?: (node: InlineNode) => InlineNode;
};

const withBookmarkIdentity = (
  marker: BookmarkStart | BookmarkEnd,
  ids: ReadonlyMap<number, PackageBookmark>,
): BookmarkStart | BookmarkEnd => {
  const alias = ids.get(marker.id);
  if (alias === undefined) return marker;
  return marker.type === "bookmarkStart"
    ? { ...marker, ...alias }
    : { ...marker, id: alias.id };
};

const withMoveIdentities = ({
  content,
  controls,
  pictures,
  bookmarkNames,
  bookmarkIds,
  reviewIdentity,
}: MoveIdentityOptions): ParagraphContent[] => {
  if (
    controls.size === 0 &&
    pictures.size === 0 &&
    !bookmarkNames?.size &&
    reviewIdentity === undefined
  )
    return content;
  const rewrite = (node: InlineNode): InlineNode => {
    const children = childNodes(node);
    const rewritten = children?.map(rewrite);
    const rebuilt =
      rewritten?.some((child, index) => child !== children?.[index]) === true
        ? rebuildNode(node, rewritten)
        : node;
    const next = reviewIdentity?.(rebuilt) ?? rebuilt;
    if (next.type === "inlineSdt" && next.properties.id !== undefined) {
      const id = controls.get(next.properties.id);
      if (id !== undefined)
        return { ...next, properties: { ...next.properties, id } };
    }
    if (next.type === "drawing") {
      const source = Number(next.image.id);
      const destination = pictures.get(source);
      if (destination !== undefined)
        return withTablePictureIdentity(next, { source, destination });
    }
    if (
      next.type === "hyperlink" &&
      next.rId === undefined &&
      next.anchor !== undefined
    ) {
      const anchor = bookmarkNames?.get(next.anchor);
      if (anchor !== undefined)
        return {
          ...next,
          anchor,
          ...(next.href === `#${next.anchor}` ? { href: `#${anchor}` } : {}),
        };
    }
    if (bookmarkIds !== undefined && bookmarkIds.size > 0) {
      if (next.type === "bookmarkStart" || next.type === "bookmarkEnd")
        return withBookmarkIdentity(next, bookmarkIds);
      if (next.type === "mathEquation") {
        const ommlXml = withCapturedBookmarkAliases(next.ommlXml, bookmarkIds);
        if (ommlXml !== next.ommlXml) return { ...next, ommlXml };
      }
    }
    return next;
  };
  return asParagraphContent(content.map(rewrite));
};
export const tablePictureIdAllocator = (
  pkg: DocxPackage,
): (() => number | undefined) => {
  const used = packageDrawingIds(pkg);
  let id = 1;
  return () => {
    while (used.has(id) && id <= MAX_TABLE_PICTURE_ID) id++;
    return id <= MAX_TABLE_PICTURE_ID ? id++ : undefined;
  };
};

export const tableBookmarkAllocator = (pkg: DocxPackage) => {
  const { ids, names } = packageBookmarkCensus(pkg);
  const occupiedNames = new Set([...names].map((name) => name.toLowerCase()));
  let nextId = 1;
  let nextName = 1;
  return (): PackageBookmark | undefined => {
    while (nextId <= MAX_REVISION_ID && ids.has(nextId)) nextId++;
    if (nextId > MAX_REVISION_ID) return undefined;
    let name = `FolioBookmark${nextName++}`;
    while (occupiedNames.has(name.toLowerCase()))
      name = `FolioBookmark${nextName++}`;
    return { id: nextId++, name };
  };
};

export const moveRangeAllocator = (
  pkg: DocxPackage,
): (() => TableMoveRanges | undefined) => {
  const { ids, names } = packageMoveRangeCensus(pkg);
  let nextId = 1;
  let nextName = 1;
  const takeId = () => {
    while (ids.has(nextId) && nextId < MAX_REVISION_ID) nextId++;
    return nextId < MAX_REVISION_ID ? nextId++ : undefined;
  };
  return () => {
    const fromRange = takeId();
    const toRange = takeId();
    if (fromRange === undefined || toRange === undefined) return undefined;
    let name = `FolioMove${nextName++}`;
    while (names.has(name)) name = `FolioMove${nextName++}`;
    return { fromRange, toRange, name };
  };
};

type MarkedTableMove = {
  record: TableGeometryMove;
  source: Paragraph;
  destination: Paragraph;
};
export const markTableMove = ({
  source,
  destination,
  from,
  to,
  mark,
  ranges,
  bookmarks,
  controls,
  pictures,
}: MarkTableMoveOptions): MarkedTableMove => {
  const { fromRange, toRange, name } = ranges;
  const aliases = bookmarkAliases(bookmarks, false);
  const record = {
    source: source.id,
    destination,
    fromRevision: from.id,
    toRevision: to.id,
    fromRange,
    toRange,
    destinationMark: mark.id,
    ...(controls.length === 0 ? {} : { controls }),
    ...(pictures.length === 0 ? {} : { pictures }),
    ...(bookmarks.length === 0 ? {} : { bookmarks }),
  } satisfies TableGeometryMove;
  return {
    record,
    source: {
      ...source.paragraph,
      content: [
        {
          type: "moveFromRangeStart",
          id: fromRange,
          name,
          author: from.author,
          date: from.date,
        },
        {
          type: "moveFrom",
          info: stampInfo(from),
          content: source.paragraph.content,
        },
        { type: "moveFromRangeEnd", id: fromRange },
      ],
    } satisfies Paragraph,
    destination: {
      ...source.paragraph,
      paraId: destination,
      pPrMark: { kind: "moveTo", info: stampInfo(mark) },
      content: [
        {
          type: "moveToRangeStart",
          id: toRange,
          name,
          author: to.author,
          date: to.date,
        },
        {
          type: "moveTo",
          info: stampInfo(to),
          content: withMoveIdentities({
            content: source.paragraph.content,
            controls: new Map(controls),
            pictures: new Map(pictures),
            bookmarkNames: aliases.bookmarkNames,
            bookmarkIds: aliases.bookmarks,
          }),
        },
        { type: "moveToRangeEnd", id: toRange },
      ],
    } satisfies Paragraph,
  };
};

export const markEmptyTableMove = ({
  source,
  destination,
  mark,
}: Pick<
  MarkTableMoveOptions,
  "source" | "destination" | "mark"
>): MarkedTableMove => ({
  record: {
    kind: "empty",
    source: source.id,
    destination,
    destinationMark: mark.id,
  } satisfies TableGeometryMove,
  source: source.paragraph,
  destination: {
    ...source.paragraph,
    paraId: destination,
    pPrMark: { kind: "ins", info: stampInfo(mark) },
  } satisfies Paragraph,
});

export const tableMoveRevisionIds = (
  moves: readonly TableGeometryMove[],
): number[] =>
  moves.flatMap((move) =>
    move.kind === "empty"
      ? [move.destinationMark]
      : [move.fromRevision, move.toRevision, move.destinationMark],
  );

type MergeGeometry = Extract<
  NonNullable<TableGridChange["geometryChange"]>,
  { type: "mergeCells" }
>;

const nestedRangeStarts = (table: Table) => {
  const starts = new Map<
    number,
    {
      type: "moveFromRangeStart" | "moveToRangeStart";
      name: string;
      paragraph: string;
    }
  >();
  for (const { paragraph } of storyParagraphs({ content: [table] })) {
    if (paragraph.paraId === undefined) continue;
    for (const start of paragraph.content) {
      if (
        start.type !== "moveFromRangeStart" &&
        start.type !== "moveToRangeStart"
      )
        continue;
      starts.set(start.id, {
        type: start.type,
        name: start.name,
        paragraph: idKey(paragraph.paraId),
      });
    }
  }
  return starts;
};

type NestedRangeOwnershipOptions = {
  source: Table;
  target: Table;
  paragraphs: readonly (readonly [string, string])[];
  ranges: NestedMoveRanges;
};
const nestedRangeOwnershipMatches = ({
  source,
  target,
  paragraphs,
  ranges,
}: NestedRangeOwnershipOptions): boolean => {
  const original = nestedRangeStarts(source);
  const copies = nestedRangeStarts(target);
  if (original.size !== copies.size) return false;
  const ids = new Map(paragraphs);
  for (const [from, to] of ranges) {
    const originalFrom = original.get(from.fromRange);
    const originalTo = original.get(from.toRange);
    const copyFrom = copies.get(to.fromRange);
    const copyTo = copies.get(to.toRange);
    if (
      originalFrom === undefined &&
      originalTo === undefined &&
      copyFrom === undefined &&
      copyTo === undefined
    )
      continue;
    if (
      originalFrom?.type !== "moveFromRangeStart" ||
      originalTo?.type !== "moveToRangeStart" ||
      copyFrom?.type !== "moveFromRangeStart" ||
      copyTo?.type !== "moveToRangeStart" ||
      originalFrom.name !== from.name ||
      originalTo.name !== from.name ||
      copyFrom.name !== to.name ||
      copyTo.name !== to.name ||
      ids.get(originalFrom.paragraph) !== copyFrom.paragraph ||
      ids.get(originalTo.paragraph) !== copyTo.paragraph
    )
      return false;
    original.delete(from.fromRange);
    original.delete(from.toRange);
    copies.delete(to.fromRange);
    copies.delete(to.toRange);
  }
  return original.size === 0 && copies.size === 0;
};

type NestedNativeMoveOptions = {
  table: Table;
  move: TableGeometryNestedMove;
  handoffs: ReadonlyMap<number, number>;
};

/** Review the inline move in both native copies while their table move is pending. */
export const nestedNativeMovePairs = ({
  table,
  move,
  handoffs,
}: NestedNativeMoveOptions): TableGeometryNativeMove[][] => {
  if (move.moveRanges === undefined) return [];
  const byRange = new Map<number, NativeHalves>();
  for (const { paragraph } of storyParagraphs({ content: [table] })) {
    if (paragraph.pPrMark !== undefined || paragraph.paraId === undefined)
      continue;
    const native = nativeMovesInParagraph(paragraph.content);
    for (const half of native?.halves ?? []) {
      const matches = byRange.get(half.range);
      if (matches === undefined) byRange.set(half.range, [half]);
      else matches.push(half);
    }
  }
  const pairs: TableGeometryNativeMove[][] = [];
  for (const [original, copy] of move.moveRanges) {
    const physicalPair = pairedNativeCellMoves(
      [
        original.fromRange,
        original.toRange,
        copy.fromRange,
        copy.toRange,
      ].flatMap((range) => byRange.get(range) ?? []),
    );
    const from = physicalPair?.find(({ name }) => name === original.name);
    const to = physicalPair?.find(({ name }) => name === copy.name);
    if (
      from !== undefined &&
      to !== undefined &&
      from.fromRange === original.fromRange &&
      from.toRange === original.toRange &&
      to.fromRange === copy.fromRange &&
      to.toRange === copy.toRange &&
      handoffs.get(from.fromRevision) === to.fromRevision &&
      handoffs.get(from.toRevision) === to.toRevision
    )
      pairs.push([from, to]);
  }
  return pairs;
};

/** Check the move's saved anchors against the native records in this merge. */
export const tableMoveOwnershipMatches = (
  table: Table,
  geometry: MergeGeometry,
): boolean => {
  if (
    geometry.moves === undefined &&
    geometry.nestedMoves === undefined &&
    geometry.emptyOwner === undefined
  )
    return true;
  const members = new Set(geometry.revisionIds);
  const sources = new Map<string, Paragraph>();
  const destinations = new Map<string, Paragraph>();
  const destinationTables = new Set<Table>();
  const movedRows = new Map<number, Table>();
  const nestedSources = new Map<string, { table: Table; ids: string[] }>();
  const unaddressedSources: Table[] = [];
  const sourceOmittedTables = new Map<number, Table>();
  const targetOmittedTables = new Map<number, Table>();
  const collectMovedTables = (
    blocks: readonly BlockContent[],
    side: "source" | "destination",
  ): void => {
    for (const block of blocks) {
      if (block.type === "table") {
        const anchor = omittedTableAnchorId(block);
        if (anchor !== undefined)
          (side === "source" ? sourceOmittedTables : targetOmittedTables).set(
            anchor,
            block,
          );
        if (side === "source") {
          const ids = paragraphIdsIn(block).map(idKey);
          if (ids.length === 0) unaddressedSources.push(block);
          const source = { table: block, ids };
          for (const id of ids) nestedSources.set(id, source);
          continue;
        }
        destinationTables.add(block);
        for (const { structuralChange } of block.rows) {
          if (structuralChange?.type === "tableRowInsertion")
            movedRows.set(structuralChange.info.id, block);
        }
      } else if (block.type === "blockCustomXml" || block.type === "blockSdt") {
        collectMovedTables(block.content, side);
      }
    }
  };
  const collect = (
    blocks: readonly BlockContent[],
    paragraphs: Map<string, Paragraph>,
  ): void => {
    for (const block of blocks) {
      if (block.type === "paragraph" && block.paraId !== undefined)
        paragraphs.set(idKey(block.paraId), block);
      else if (block.type === "blockCustomXml" || block.type === "blockSdt")
        collect(block.content, paragraphs);
    }
  };
  let hasOwner = false;
  let emptyOwner: string | undefined;
  for (const row of table.rows) {
    const removedRow =
      row.structuralChange?.type === "tableRowDeletion" &&
      members.has(row.structuralChange.info.id);
    for (const cell of row.cells) {
      const removedCell =
        cell.structuralChange?.type === "tableCellDeletion" &&
        members.has(cell.structuralChange.info.id);
      const owned =
        cell.propertyChanges?.some(({ info }) => members.has(info.id)) ?? false;
      if (!owned && !removedRow && !removedCell) continue;
      if (owned) {
        if (hasOwner || removedRow || removedCell) return false;
        hasOwner = true;
      }
      const first = cell.content.at(0);
      if (owned && first?.type === "paragraph") emptyOwner = first.paraId;
      if (geometry.nestedMoves !== undefined)
        collectMovedTables(cell.content, owned ? "destination" : "source");
      collect(cell.content, owned ? destinations : sources);
    }
  }
  if (
    geometry.emptyOwner !== undefined &&
    (emptyOwner === undefined ||
      idKey(geometry.emptyOwner) !== idKey(emptyOwner) ||
      (geometry.moves?.length ?? 0) + (geometry.nestedMoves?.length ?? 0) === 0)
  )
    return false;
  const nestedTables = new Set<Table>();
  const omittedSources = new Set<Table>();
  const claimedSources = new Set<Table>();
  const surviving = new Set(
    geometry.nestedMoves === undefined ? [] : paragraphIdsIn(table).map(idKey),
  );
  for (const move of geometry.nestedMoves ?? []) {
    if (move.omittedTable !== undefined) {
      const [from, to] = move.omittedTable;
      if (from === to || move.rows.length !== 0 || move.paragraphs.length !== 0)
        return false;
      const source = sourceOmittedTables.get(from);
      const target = targetOmittedTables.get(to);
      if (source === undefined && target === undefined) {
        // Reject misplaced records, while allowing an earlier row decision to
        // leave both copies empty and remove them completely.
        const physical = new Set(identityKeysIn(table));
        if (
          [from, to].some((id) =>
            physical.has(slotKey({ space: IDENTITY_SPACES.REVISION, id })),
          )
        )
          return false;
        continue;
      }
      if (
        source === undefined ||
        target === undefined ||
        omittedSources.has(source) ||
        nestedTables.has(target) ||
        !bookmarkOwnershipMatches(source, target, move.bookmarks)
      )
        return false;
      omittedSources.add(source);
      nestedTables.add(target);
      continue;
    }
    const first = move.rows.at(0);
    const destination =
      first === undefined
        ? nestedTableCopyByAnchors(table, move.paragraphs)
        : movedRows.get(first);
    // An earlier review can remove paired paragraphs from both copies.
    const pairs = move.paragraphs
      .map(([source, target]) => [idKey(source), idKey(target)] as const)
      .filter(
        ([source, target]) => surviving.has(source) || surviving.has(target),
      );
    if (
      destination === undefined &&
      first === undefined &&
      move.paragraphs.length > 0 &&
      pairs.length === 0
    )
      continue;
    if (
      destination === undefined ||
      !destinationTables.has(destination) ||
      nestedTables.has(destination) ||
      new Set(move.rows).size !== move.rows.length ||
      move.rows.some((id) => movedRows.get(id) !== destination)
    )
      return false;
    const destinationIds = paragraphIdsIn(destination).map(idKey);
    if (destinationIds.length !== pairs.length) return false;
    const firstSource = pairs.at(0)?.at(0);
    if (firstSource !== undefined) {
      const source = nestedSources.get(firstSource);
      if (
        source?.ids.length !== pairs.length ||
        claimedSources.has(source.table) ||
        pairs.some(
          ([from, to], index) =>
            source.ids[index] !== from || destinationIds[index] !== to,
        ) ||
        !bookmarkOwnershipMatches(source.table, destination, move.bookmarks) ||
        !nestedRangeOwnershipMatches({
          source: source.table,
          target: destination,
          paragraphs: pairs,
          ranges: move.moveRanges ?? [],
        })
      )
        return false;
      claimedSources.add(source.table);
    } else if (move.bookmarks !== undefined) {
      const matchingSources = unaddressedSources.filter(
        (source) =>
          !claimedSources.has(source) &&
          bookmarkOwnershipMatches(source, destination, move.bookmarks),
      );
      const source = matchingSources.at(0);
      if (matchingSources.length !== 1 || source === undefined) return false;
      claimedSources.add(source);
    }
    nestedTables.add(destination);
  }
  const seen = new Set<string>();
  for (const move of geometry.moves ?? []) {
    const sourceId = idKey(move.source);
    const destinationId = idKey(move.destination);
    if (
      sourceId === destinationId ||
      seen.has(sourceId) ||
      seen.has(destinationId)
    )
      return false;
    seen.add(sourceId);
    seen.add(destinationId);
    const source = sources.get(sourceId);
    const destination = destinations.get(destinationId);
    if (move.kind === "empty") {
      if (
        source?.content.length !== 0 ||
        destination?.content.length !== 0 ||
        destination.pPrMark?.kind !== "ins" ||
        destination.pPrMark.info.id !== move.destinationMark
      )
        return false;
      continue;
    }
    if (
      source?.content.length !== 3 ||
      destination?.content.length !== 3 ||
      destination.pPrMark?.kind !== "moveTo" ||
      destination.pPrMark.info.id !== move.destinationMark
    )
      return false;
    const [fromStart, from, fromEnd] = source.content;
    const [toStart, to, toEnd] = destination.content;
    if (
      fromStart?.type !== "moveFromRangeStart" ||
      fromStart.id !== move.fromRange ||
      from?.type !== "moveFrom" ||
      from.info.id !== move.fromRevision ||
      fromEnd?.type !== "moveFromRangeEnd" ||
      fromEnd.id !== move.fromRange ||
      toStart?.type !== "moveToRangeStart" ||
      toStart.id !== move.toRange ||
      to?.type !== "moveTo" ||
      to.info.id !== move.toRevision ||
      toEnd?.type !== "moveToRangeEnd" ||
      toEnd.id !== move.toRange ||
      fromStart.name !== toStart.name ||
      (move.bookmarks !== undefined &&
        !bookmarkOwnershipMatches(
          { ...source, content: from.content },
          { ...destination, content: to.content },
          move.bookmarks,
        ))
    )
      return false;
  }
  return true;
};

/** The control shells on opposite sides of the recorded paragraph moves. */
export const tableBlockControlHandoffs = (
  table: Table,
): ReadonlyMap<number, number> => {
  const controls = new Map<number, number>();
  const geometry = table.formatting?.gridChange?.geometryChange;
  if (geometry?.type !== "mergeCells") return controls;
  const pairs = [
    ...(geometry.moves ?? []).map(
      ({ source, destination }) => [source, destination] as const,
    ),
    ...(geometry.nestedMoves ?? []).flatMap(({ paragraphs }) => paragraphs),
  ];
  if (pairs.length === 0) return controls;
  const anchors = new Map<string, readonly BlockSdt[]>();
  const collect = (
    blocks: readonly BlockContent[],
    scope: readonly BlockSdt[],
  ): void => {
    for (const block of blocks) {
      if (block.type === "blockSdt") collect(block.content, [...scope, block]);
      else if (block.type === "blockCustomXml") collect(block.content, scope);
      else if (block.type === "table") {
        for (const row of block.rows)
          for (const cell of row.cells) collect(cell.content, scope);
      } else if (
        block.type === "paragraph" &&
        block.paraId !== undefined &&
        scope.length > 0
      ) {
        anchors.set(idKey(block.paraId), scope);
      }
    }
  };
  for (const row of table.rows)
    for (const cell of row.cells) collect(cell.content, []);
  for (const [source, destination] of pairs) {
    const originals = anchors.get(idKey(source));
    const copies = anchors.get(idKey(destination));
    if (
      originals === undefined ||
      copies === undefined ||
      originals.length !== copies.length
    )
      continue;
    for (const [index, original] of originals.entries()) {
      const copy = copies[index];
      if (
        copy === undefined ||
        !movableBlockControl(original) ||
        !movableBlockControl(copy)
      )
        continue;
      const { id: from, ...originalProperties } = original.properties;
      const { id: to, ...copyProperties } = copy.properties;
      if (
        from !== undefined &&
        to !== undefined &&
        from !== to &&
        structurallyEqual(originalProperties, copyProperties)
      )
        controls.set(to, from);
    }
  }
  return controls;
};

type ResolveTableMovesOptions = {
  table: Table;
  reject: boolean;
  handoffs: TableMoveHandoffs;
};

export type TableMoveHandoffs = {
  revisions: ReadonlyMap<number, number>;
  controls: ReadonlyMap<number, number>;
};

const withoutOmittedTableAnchor = (table: Table): Table => {
  const { propertyChanges, ...base } = table;
  const remaining = propertyChanges?.slice(1);
  return remaining === undefined || remaining.length === 0
    ? base
    : { ...base, propertyChanges: remaining };
};

/** Settle copied identities with the cell topology. */
export const resolveTableMoves = ({
  table,
  reject,
  handoffs,
}: ResolveTableMovesOptions): Table => {
  const geometry = table.formatting?.gridChange?.geometryChange;
  if (
    geometry?.type !== "mergeCells" ||
    (geometry.moves === undefined && geometry.nestedMoves === undefined)
  )
    return table;
  const destinations = new Map(
    (geometry.moves ?? []).map((move) => [idKey(move.destination), move]),
  );
  const sources = new Map(
    (geometry.moves ?? []).map((move) => [idKey(move.source), move]),
  );
  const nested = new Map(
    (geometry.nestedMoves ?? []).flatMap((move) =>
      move.rows.map((id) => [id, move] as const),
    ),
  );
  const nestedAnchors = new Map<Table, TableGeometryNestedMove>();
  const omittedSources = new Set<number>();
  const omittedTargets = new Map<number, TableGeometryNestedMove>();
  for (const move of geometry.nestedMoves ?? []) {
    if (move.omittedTable !== undefined) {
      omittedSources.add(move.omittedTable[0]);
      omittedTargets.set(move.omittedTable[1], move);
      continue;
    }
    if (move.rows.length > 0) continue;
    const copy = nestedTableCopyByAnchors(table, move.paragraphs);
    if (copy !== undefined) nestedAnchors.set(copy, move);
  }
  const emptyOwner =
    geometry.emptyOwner === undefined ? undefined : idKey(geometry.emptyOwner);
  const liveParagraphs = new Set(paragraphIdsIn(table).map(idKey));
  const nestedMoves = geometry.nestedMoves ?? [];
  const liveNestedRowRevisions = new Set<number>();
  const collectNestedRowRevisions = (blocks: readonly BlockContent[]): void => {
    for (const block of blocks) {
      if (block.type === "blockCustomXml" || block.type === "blockSdt") {
        collectNestedRowRevisions(block.content);
        continue;
      }
      if (block.type !== "table") continue;
      for (const row of block.rows) {
        if (row.structuralChange !== undefined)
          liveNestedRowRevisions.add(row.structuralChange.info.id);
        for (const cell of row.cells) collectNestedRowRevisions(cell.content);
      }
    }
  };
  for (const row of table.rows)
    for (const cell of row.cells) collectNestedRowRevisions(cell.content);
  const substantiveDestinationIds = [
    ...(geometry.moves ?? []).flatMap((move) =>
      move.kind === "empty" ? [] : [idKey(move.destination)],
    ),
    ...nestedMoves.flatMap((move) =>
      move.paragraphs.map(([, destination]) => idKey(destination)),
    ),
  ];
  // Empty paragraphs copied around a nested table are structural move
  // boundaries. If earlier accepted edits removed every substantive copied
  // block, the direct merge sees an empty donor and copies none of them.
  const discardEmptyOnlyMoveDestinations =
    !reject &&
    nestedMoves.length > 0 &&
    nestedMoves.every(
      (move) =>
        move.omittedTable === undefined &&
        move.rows.length > 0 &&
        move.paragraphs.length > 0,
    ) &&
    !substantiveDestinationIds.some((id) => liveParagraphs.has(id)) &&
    !nestedMoves.some(
      (move) =>
        move.rows.some((id) => liveNestedRowRevisions.has(id)) ||
        move.revisions?.some(
          ([source, destination]) =>
            liveNestedRowRevisions.has(source) ||
            liveNestedRowRevisions.has(destination),
        ),
    );
  const rows = table.rows.map((row) =>
    Object.assign({}, row, {
      cells: row.cells.map((cell) =>
        Object.assign({}, cell, {
          content: cell.content.flatMap(
            function restore(block): BlockContent[] {
              if (
                block.type === "blockCustomXml" ||
                block.type === "blockSdt"
              ) {
                const content = block.content.flatMap(restore);
                // A destination wrapper whose copied children are all gone
                // goes with them. Keep pre-existing empty wrappers.
                if (content.length === 0 && block.content.length > 0) return [];
                const id =
                  block.type === "blockSdt" && block.properties.id !== undefined
                    ? handoffs.controls.get(block.properties.id)
                    : undefined;
                if (!reject && block.type === "blockSdt" && id !== undefined)
                  return [
                    {
                      ...block,
                      properties: { ...block.properties, id },
                      content,
                    },
                  ];
                return [{ ...block, content }];
              }
              if (block.type === "table") {
                const anchor =
                  omittedSources.size === 0
                    ? undefined
                    : omittedTableAnchorId(block);
                if (
                  reject &&
                  anchor !== undefined &&
                  omittedSources.has(anchor)
                )
                  return [withoutOmittedTableAnchor(block)];
                const omitted =
                  anchor === undefined ? undefined : omittedTargets.get(anchor);
                const move = block.rows
                  .map((nestedRow) => nestedRow.structuralChange)
                  .find(
                    (change) =>
                      change?.type === "tableRowInsertion" &&
                      nested.has(change.info.id),
                  );
                const record =
                  omitted ??
                  (move === undefined
                    ? nestedAnchors.get(block)
                    : nested.get(move.info.id));
                if (record === undefined) return [block];
                if (reject) return [];
                const placed = withNestedMoveIdentities(block, {
                  paragraphs: new Map(
                    record.paragraphs.map(([source, destination]) => [
                      idKey(destination),
                      source,
                    ]),
                  ),
                  revisions: handoffs.revisions,
                  controls: new Map(
                    record.controls?.map(([old, temporary]) => [
                      temporary,
                      old,
                    ]),
                  ),
                  pictures: new Map(
                    record.pictures?.map(([old, temporary]) => [
                      temporary,
                      old,
                    ]),
                  ),
                  ranges: nestedMoveRangeMap(record.moveRanges ?? [], true),
                  ...bookmarkAliases(record.bookmarks ?? [], true),
                });
                const stamps = new Set(record.rows);
                placed.rows = placed.rows.map((child) => {
                  const change = child.structuralChange;
                  if (change === undefined || !stamps.has(change.info.id))
                    return child;
                  const { structuralChange: _owned, ...accepted } = child;
                  return accepted;
                });
                return [
                  omitted === undefined
                    ? placed
                    : withoutOmittedTableAnchor(placed),
                ];
              }
              if (block.type !== "paragraph" || block.paraId === undefined)
                return [block];
              const key = idKey(block.paraId);
              const destination = destinations.get(key);
              if (
                discardEmptyOnlyMoveDestinations &&
                destination?.kind === "empty"
              )
                return [];
              if (
                (reject && destination !== undefined) ||
                (!reject &&
                  !discardEmptyOnlyMoveDestinations &&
                  key === emptyOwner &&
                  block.content.length === 0 &&
                  block.pPrMark === undefined &&
                  block.propertyChanges === undefined &&
                  block.sectionProperties === undefined)
              )
                return [];
              const move = reject ? sources.get(key) : destination;
              if (move === undefined) return [block];
              let revision: number | undefined;
              let range: number | undefined;
              if (move.kind !== "empty") {
                revision = reject ? move.fromRevision : move.toRevision;
                range = reject ? move.fromRange : move.toRange;
              }
              const content = block.content.flatMap(
                (item): ParagraphContent[] => {
                  switch (item.type) {
                    case "moveFromRangeStart":
                    case "moveFromRangeEnd":
                    case "moveToRangeStart":
                    case "moveToRangeEnd":
                      return item.id === range ? [] : [item];
                    case "moveFrom":
                    case "moveTo":
                      return item.info.id === revision ? item.content : [item];
                    default:
                      return [item];
                  }
                },
              );
              let restored = content;
              if (!reject && move.kind !== "empty") {
                const aliases = bookmarkAliases(move.bookmarks ?? [], true);
                restored = withMoveIdentities({
                  content,
                  controls: new Map(
                    move.controls?.map(([old, temporary]) => [temporary, old]),
                  ),
                  pictures: new Map(
                    move.pictures?.map(([old, temporary]) => [temporary, old]),
                  ),
                  bookmarkNames: aliases.bookmarkNames,
                  bookmarkIds: aliases.bookmarks,
                });
              }
              const paragraph = {
                ...block,
                paraId: move.source,
                content: restored,
              };
              if (paragraph.pPrMark?.info.id === move.destinationMark)
                delete paragraph.pPrMark;
              return [paragraph];
            },
          ),
        }),
      ),
    }),
  );
  return { ...table, rows };
};
