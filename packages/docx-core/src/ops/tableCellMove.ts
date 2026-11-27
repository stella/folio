import { isNoteReferenceId } from "../model/content";
import {
  type BlockContent,
  type BlockCustomXml,
  type BlockSdt,
  type DocxPackage,
  type InlineSdt,
  type MathEquation,
  MAX_REVISION_ID,
  type Paragraph,
  type ParagraphContent,
  type PreservedInline,
  type PreservedXmlContent,
  type Run,
  type RunContent,
  type SdtProperties,
  type Table,
  type TableCell,
  type TableRow,
} from "../model/document";
import {
  MAX_TABLE_PICTURE_ID,
  type TableGeometryMove,
  type TableGeometryNestedMove,
  type TableGridChange,
} from "../model/formatting";
import { nestedTableCopyByAnchors } from "../model/nestedTableRevisionHandoffs";
import { readXml, type XmlElement } from "../xml/streamingXmlParser";
import {
  OOXML_NAMESPACE_SCOPE,
  OOXML_NS,
  resolveNamespaceUri,
  XML_NAMESPACE_URI,
} from "../xml/xmlNamespaceContext";
import { blockListEndsInParagraph } from "./blocks";
import { structurallyEqual } from "./equality";
import {
  IDENTITY_SPACES,
  identityKeysIn,
  idKey,
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
import {
  movableTablePictureId,
  withTablePictureIdentity,
} from "./tablePictureMove";
import type { RevisionStamp } from "./types";

export type TableMoveSource = {
  paragraph: Paragraph;
  id: string;
  controls: number[];
  pictures: ReadonlySet<number>;
};
export type NestedTableMoveSource = {
  table: Table;
  paragraphs: string[];
  revisions: number[];
  insertedRows: number;
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
  run.propertyChanges === undefined &&
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

const localMathNode = (node: XmlElement): boolean => {
  if (node.type === "text" || node.type === "comment") return true;
  if (node.type !== "element" || !mathNamespace(node.namespaceUri))
    return false;
  return (
    Object.keys(node.attributes ?? {}).every((name) => {
      if (name === "xmlns" || name.startsWith("xmlns:")) return true;
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
    (node.elements?.every(localMathNode) ?? true)
  );
};

/** A native equation can appear in both move halves if it owns no other markup. */
const movableMath = ({ display, ommlXml }: MathEquation): boolean => {
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
    localMathNode(root)
  );
};

const emptyElement = (xml: string) => {
  const parsed = readXml(xml, {
    conformance: "ooxml",
    inheritedNamespaceScope: OOXML_NAMESPACE_SCOPE,
  });
  if (parsed.status !== "parsed" || parsed.value.elements?.length !== 1)
    return undefined;
  const element = parsed.value.elements.at(0);
  if (element?.type !== "element" || (element.elements?.length ?? 0) > 0)
    return undefined;
  const attributes = Object.entries(element.attributes ?? {}).filter(
    ([name]) => name !== "xmlns" && !name.startsWith("xmlns:"),
  );
  const localName = element.name?.slice(element.name.indexOf(":") + 1);
  return { element, localName, attributes };
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

const controlPropertyKind = (
  xml: string,
): "richText" | "binding" | undefined => {
  const parsed = emptyElement(xml);
  if (parsed === undefined) return undefined;
  const { element, localName, attributes } = parsed;
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
  if (
    properties.sdtType !== "richText" ||
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
    if (kind === undefined || kinds.has(kind)) return undefined;
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
        return movableRun(item, context.pictures);
      case "mathEquation":
        return movableMath(item);
      case "preservedInline":
        return recordProofingMarker(item, context.proofing);
      case "hyperlink":
        // A bidi wrapper can contain a link, but cannot hide a nested one.
        return (
          !context.withinLink &&
          movableInlineContent(item.children, { ...context, withinLink: true })
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
): TableMoveSource | undefined => {
  const controls: number[] = [];
  const proofing = new Set<ProofingKind>();
  const pictures = new Set<number>();
  if (
    paragraph.paraId === undefined ||
    paragraph.pPrMark !== undefined ||
    paragraph.propertyChanges !== undefined ||
    paragraph.sectionProperties !== undefined ||
    paragraph.reviewCarrier !== undefined ||
    (paragraph.content.length !== 0 &&
      !movableInlineContent(paragraph.content, {
        withinLink: false,
        controls,
        proofing,
        pictures,
      })) ||
    proofing.size !== 0
  )
    return undefined;
  return { paragraph, id: paragraph.paraId, controls, pictures };
};

/** A structural copy may carry only records whose package identities stay unique. */
export const nestedTableMoveSource = (
  table: Table,
): NestedTableMoveSource | undefined => {
  const paragraphs: string[] = [];
  const revisions: number[] = [];
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
  }: TableRow | TableCell): boolean =>
    (contentControls?.every(collectControl) ?? true) &&
    (carrierStack?.every((carrier) =>
      carrier.type === "customXml"
        ? movableCustomXmlBlock(carrier.wrapper)
        : collectControl(carrier.properties),
    ) ??
      true);
  const collectProperties = (
    changes: readonly { info: { id: number } }[] | undefined,
  ) => {
    for (const { info } of changes ?? []) revisions.push(info.id);
  };
  const collect = ({ rows, ...tableProperties }: Table): boolean => {
    const grid = tableProperties.formatting?.gridChange;
    const geometry = grid?.geometryChange;
    if (
      tableProperties.bookmarks !== undefined ||
      tableProperties.carrierStack !== undefined ||
      reachesPastRemovedCell(tableProperties) ||
      (geometry?.type === "mergeCells" &&
        (geometry.moves !== undefined ||
          geometry.nestedMoves !== undefined ||
          geometry.emptyOwner !== undefined)) ||
      rows.length === 0
    )
      return false;
    if (grid !== undefined) {
      revisions.push(grid.id);
      for (const id of geometry?.revisionIds ?? []) gridMembers.push(id);
    }
    collectProperties(tableProperties.propertyChanges);
    return rows.every((row) => {
      const {
        cells,
        contentControls: _rowControls,
        carrierStack: _rowStack,
        ...rowProperties
      } = row;
      if (
        rowProperties.bookmarks !== undefined ||
        !collectCarriers(row) ||
        reachesPastRemovedCell(rowProperties) ||
        cells.length === 0
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
          cell.propertyChanges?.some(
            ({ previousStructuralChange }) =>
              previousStructuralChange !== undefined,
          ) ||
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
        return content.every(collectBlock);
      });
    });
  };
  const collectBlock = (block: BlockContent): boolean => {
    if (block.type === "table") return collect(block);
    if (block.type === "blockCustomXml")
      return movableCustomXmlBlock(block) && block.content.every(collectBlock);
    if (block.type === "blockSdt") {
      return (
        collectControl(block.properties) && block.content.every(collectBlock)
      );
    }
    if (block.type !== "paragraph") return false;
    const source = tableMoveSource(block);
    if (source === undefined) return false;
    paragraphs.push(source.id);
    for (const id of source.controls) controls.add(id);
    for (const id of source.pictures) {
      if (pictures.has(id)) return false;
      pictures.add(id);
    }
    return true;
  };
  if (!collect(table)) return undefined;
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
        controls: [...controls],
        pictures,
        insertedRows: table.rows.filter(
          ({ structuralChange }) => structuralChange === undefined,
        ).length,
      }
    : undefined;
};

type NestedMoveIdentities = {
  paragraphs: ReadonlyMap<string, string>;
  revisions: ReadonlyMap<number, number>;
  controls: ReadonlyMap<number, number>;
  pictures: ReadonlyMap<number, number>;
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
  const carrierIdentities = <T extends TableRow | TableCell>(item: T): T => ({
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
  const restamp = <Change extends { info: { id: number } }>(
    record: Change,
  ): Change => {
    const id = ids.revisions.get(record.info.id);
    return id === undefined
      ? record
      : { ...record, info: { ...record.info, id } };
  };
  const rewrite = (block: BlockContent): BlockContent => {
    if (block.type === "table") return withNestedMoveIdentities(block, ids);
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
          content: withMoveIdentities({
            content: block.content,
            controls: ids.controls,
            pictures: ids.pictures,
          }),
        };
  };
  const grid = table.formatting?.gridChange;
  const gridId = grid === undefined ? undefined : ids.revisions.get(grid.id);
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
              ...(grid.geometryChange !== undefined
                ? {
                    geometryChange: {
                      ...grid.geometryChange,
                      revisionIds: grid.geometryChange.revisionIds.map(
                        (id) => ids.revisions.get(id) ?? id,
                      ),
                    },
                  }
                : {}),
            },
          },
        }),
    ...(table.propertyChanges === undefined
      ? {}
      : { propertyChanges: table.propertyChanges.map(restamp) }),
    rows: table.rows.map((row) => ({
      ...carrierIdentities(row),
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
          : { propertyChanges: cell.propertyChanges.map(restamp) }),
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
  controls: readonly (readonly [number, number])[];
  pictures: readonly (readonly [number, number])[];
};
type MarkedNestedTableMove = {
  destination: Table;
  record: TableGeometryNestedMove;
};
export const markNestedTableMove = ({
  source,
  destinations,
  rowStamps,
  recordStamps,
  controls,
  pictures,
}: MarkNestedTableMoveOptions): MarkedNestedTableMove | undefined => {
  if (
    destinations.length !== source.paragraphs.length ||
    rowStamps.length !== source.insertedRows ||
    recordStamps.length !== source.revisions.length
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
  return {
    destination: { ...table, rows },
    record: {
      rows: rowStamps.map(({ id }) => id),
      paragraphs,
      ...(revisions.length === 0 ? {} : { revisions }),
      ...(controls.length === 0 ? {} : { controls }),
      ...(pictures.length === 0 ? {} : { pictures }),
    } satisfies TableGeometryNestedMove,
  };
};

type TableMoveRanges = { fromRange: number; toRange: number; name: string };
type MarkTableMoveOptions = {
  source: TableMoveSource;
  destination: string;
  from: RevisionStamp;
  to: RevisionStamp;
  mark: RevisionStamp;
  ranges: TableMoveRanges;
  controls: readonly (readonly [number, number])[];
  pictures: readonly (readonly [number, number])[];
};

type MoveIdentityOptions = {
  content: ParagraphContent[];
  controls: ReadonlyMap<number, number>;
  pictures: ReadonlyMap<number, number>;
};

const withMoveIdentities = ({
  content,
  controls,
  pictures,
}: MoveIdentityOptions): ParagraphContent[] => {
  if (controls.size === 0 && pictures.size === 0) return content;
  const rewrite = (node: InlineNode): InlineNode => {
    const children = childNodes(node);
    const rewritten = children?.map(rewrite);
    const next =
      rewritten?.some((child, index) => child !== children?.[index]) === true
        ? rebuildNode(node, rewritten)
        : node;
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
  controls,
  pictures,
}: MarkTableMoveOptions): MarkedTableMove => {
  const { fromRange, toRange, name } = ranges;
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
  const movedRows = new Map<number, Table>();
  const nestedSources = new Map<string, string[]>();
  const collectMovedTables = (
    blocks: readonly BlockContent[],
    side: "source" | "destination",
  ): void => {
    for (const block of blocks) {
      if (block.type === "table") {
        if (side === "source") {
          const ids = paragraphIdsIn(block).map(idKey);
          const first = ids.at(0);
          if (first !== undefined) nestedSources.set(first, ids);
          continue;
        }
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
  const claimedSources = new Set<string>();
  for (const move of geometry.nestedMoves ?? []) {
    const first = move.rows.at(0);
    const destination = first === undefined ? undefined : movedRows.get(first);
    if (
      destination === undefined ||
      nestedTables.has(destination) ||
      new Set(move.rows).size !== move.rows.length ||
      move.rows.some((id) => movedRows.get(id) !== destination)
    )
      return false;
    const firstSource = move.paragraphs.at(0)?.at(0);
    if (firstSource === undefined) return false;
    const sourceId = idKey(firstSource);
    const sourceIds = nestedSources.get(sourceId);
    const destinationIds = paragraphIdsIn(destination).map(idKey);
    if (
      claimedSources.has(sourceId) ||
      sourceIds?.length !== move.paragraphs.length ||
      destinationIds.length !== move.paragraphs.length ||
      move.paragraphs.some(
        ([from, to], index) =>
          sourceIds[index] !== idKey(from) ||
          destinationIds[index] !== idKey(to),
      )
    )
      return false;
    claimedSources.add(sourceId);
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
      fromStart.name !== toStart.name
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
  for (const move of geometry.nestedMoves ?? []) {
    if (move.rows.length > 0) continue;
    const copy = nestedTableCopyByAnchors(table, move.paragraphs);
    if (copy !== undefined) nestedAnchors.set(copy, move);
  }
  const emptyOwner =
    geometry.emptyOwner === undefined ? undefined : idKey(geometry.emptyOwner);
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
                const move = block.rows
                  .map((nestedRow) => nestedRow.structuralChange)
                  .find(
                    (change) =>
                      change?.type === "tableRowInsertion" &&
                      nested.has(change.info.id),
                  );
                const record =
                  move === undefined
                    ? nestedAnchors.get(block)
                    : nested.get(move.info.id);
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
                });
                const stamps = new Set(record.rows);
                placed.rows = placed.rows.map((child) => {
                  const change = child.structuralChange;
                  if (change === undefined || !stamps.has(change.info.id))
                    return child;
                  const { structuralChange: _owned, ...accepted } = child;
                  return accepted;
                });
                return [placed];
              }
              if (block.type !== "paragraph" || block.paraId === undefined)
                return [block];
              const key = idKey(block.paraId);
              const destination = destinations.get(key);
              if (
                (reject && destination !== undefined) ||
                (!reject &&
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
              const restored =
                reject ||
                move.kind === "empty" ||
                (move.controls === undefined && move.pictures === undefined)
                  ? content
                  : withMoveIdentities({
                      content,
                      controls: new Map(
                        move.controls?.map(([old, temporary]) => [
                          temporary,
                          old,
                        ]),
                      ),
                      pictures: new Map(
                        move.pictures?.map(([old, temporary]) => [
                          temporary,
                          old,
                        ]),
                      ),
                    });
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
