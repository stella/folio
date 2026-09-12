import { panic } from "better-result";
import type { Node as PMNode } from "prosemirror-model";

import { PROSE_PARAGRAPH_SOURCE_TOKEN_ATTR } from "../../docx/paragraphPropertySource";
import {
  paragraphAttrsToFormatting,
  tableAttrsToFormatting,
  tableCellAttrsToFormatting,
  tableRowAttrsToFormatting,
} from "../../prosemirror/conversion/fromProseDoc";
import { listRenderingFromAttrs } from "../../prosemirror/listRenderingProjection";
import type { ParagraphProjectionOnlyAttrName } from "../../prosemirror/paragraphProjectionAttrs";
import {
  expectParagraphAttrs,
  expectTableAttrs,
  expectTableCellAttrs,
  expectTableRowAttrs,
} from "../../prosemirror/attrs";
import type {
  ParagraphAttrs,
  TableAttrs,
  TableCellAttrs,
  TableRowAttrs,
  TextBoxAnchorAttrs,
  TextBoxAttrs,
} from "../../prosemirror/schema/nodes";
import { canonicalJson } from "../../utils/canonicalJson";

/**
 * `semantic` values serialize exactly; `transport` values bind the live tree
 * to its package source or preserve opaque package state; `presentation`
 * values are rebuildable editor caches. The remaining dispositions name the
 * canonical serializer projection that owns a cluster of dependent attrs.
 */
type SourceIdentityAttrDisposition =
  | "semantic"
  | "transport"
  | "presentation"
  | "paragraph-document"
  | "paragraph-section"
  | "paragraph-bookmark-boundaries"
  | "table-document"
  | "table-row-document"
  | "table-cell-document"
  | "text-box-anchor"
  | "text-box-group";

type ParagraphSourceIdentityAttrName =
  | keyof ParagraphAttrs
  | ParagraphProjectionOnlyAttrName
  | typeof PROSE_PARAGRAPH_SOURCE_TOKEN_ATTR;

/**
 * Total source-identity ownership for the paragraph schema. Formatting and
 * list attrs compare through the serializer's canonical Document projection;
 * presentation-only caches cannot masquerade as serialized differences.
 */
const PARAGRAPH_SOURCE_IDENTITY_ATTRS = Object.freeze({
  paraId: "transport",
  textId: "transport",
  alignment: "paragraph-document",
  alignmentFromStyle: "paragraph-document",
  kinsoku: "paragraph-document",
  overflowPunctuation: "paragraph-document",
  suppressAutoHyphens: "paragraph-document",
  spaceBefore: "paragraph-document",
  spaceAfter: "paragraph-document",
  lineSpacing: "paragraph-document",
  lineSpacingRule: "paragraph-document",
  lineSpacingExplicit: "paragraph-document",
  snapToGrid: "paragraph-document",
  spacingExplicit: "paragraph-document",
  spacingFromDocDefaults: "paragraph-document",
  spacingFromImplicitDefaultStyle: "paragraph-document",
  indentLeft: "paragraph-document",
  indentRight: "paragraph-document",
  indentFirstLine: "paragraph-document",
  hangingIndent: "paragraph-document",
  numPr: "paragraph-document",
  numPrFromStyle: "paragraph-document",
  listNumFmt: "paragraph-document",
  listIsBullet: "paragraph-document",
  listIsLegal: "paragraph-document",
  listMarker: "paragraph-document",
  listMarkerTemplate: "paragraph-document",
  listMarkerHidden: "paragraph-document",
  listMarkerFormatting: "paragraph-document",
  listMarkerAlignment: "paragraph-document",
  listMarkerSuffix: "paragraph-document",
  listMarkerAllCaps: "paragraph-document",
  listImplicitChildLevelAdvances: "paragraph-document",
  listMarkerSecondSlotOffsetTwips: "paragraph-document",
  listLevelNumFmts: "paragraph-document",
  listLevelStarts: "paragraph-document",
  listAbstractNumId: "paragraph-document",
  listStartOverride: "paragraph-document",
  styleId: "paragraph-document",
  _tableOfContentsLevel: "presentation",
  borders: "paragraph-document",
  shading: "paragraph-document",
  tabs: "paragraph-document",
  pageBreakBefore: "paragraph-document",
  renderedPageBreakBefore: "transport",
  _pageBreakCarrier: "presentation",
  _trailingPageBreak: "presentation",
  keepNext: "paragraph-document",
  keepLines: "paragraph-document",
  widowControl: "paragraph-document",
  contextualSpacing: "paragraph-document",
  defaultTextFormatting: "semantic",
  sectionBreakType: "paragraph-section",
  direction: "paragraph-document",
  outlineLevel: "paragraph-document",
  bookmarks: "paragraph-bookmark-boundaries",
  _emptyHyperlinks: "semantic",
  runInWithNext: "paragraph-document",
  _originalFormatting: "paragraph-document",
  _autospacingBase: "paragraph-document",
  _sectionProperties: "paragraph-section",
  _propertyChanges: "semantic",
  pPrMark: "semantic",
  _suggestedInsert: "semantic",
  idStability: "transport",
  _detachedWatermarkHost: "presentation",
  [PROSE_PARAGRAPH_SOURCE_TOKEN_ATTR]: "transport",
} as const satisfies Readonly<
  Record<ParagraphSourceIdentityAttrName, SourceIdentityAttrDisposition>
>);

const TABLE_SOURCE_IDENTITY_ATTRS = Object.freeze({
  styleId: "table-document",
  width: "table-document",
  widthType: "table-document",
  justification: "table-document",
  columnWidths: "semantic",
  floating: "table-document",
  cellMargins: "table-document",
  look: "table-document",
  borders: "table-document",
  _resolvedCellMargins: "table-document",
  _resolvedIndent: "presentation",
  _resolvedJustification: "presentation",
  _resolvedBidi: "presentation",
  _originalFormatting: "table-document",
  tblPrChange: "semantic",
  _suggestedInsert: "semantic",
} as const satisfies Readonly<Record<keyof TableAttrs, SourceIdentityAttrDisposition>>);

const TABLE_ROW_SOURCE_IDENTITY_ATTRS = Object.freeze({
  height: "table-row-document",
  heightRule: "table-row-document",
  isHeader: "table-row-document",
  hidden: "table-row-document",
  _resolvedJustification: "presentation",
  _originalFormatting: "table-row-document",
  trPrChange: "semantic",
  trIns: "semantic",
  trDel: "semantic",
} as const satisfies Readonly<Record<keyof TableRowAttrs, SourceIdentityAttrDisposition>>);

const TABLE_CELL_SOURCE_IDENTITY_ATTRS = Object.freeze({
  colspan: "table-cell-document",
  rowspan: "table-cell-document",
  colwidth: "presentation",
  width: "table-cell-document",
  widthType: "table-cell-document",
  verticalAlign: "table-cell-document",
  backgroundColor: "table-cell-document",
  _resolvedBackgroundColor: "table-cell-document",
  textDirection: "table-cell-document",
  noWrap: "table-cell-document",
  hideMark: "presentation",
  borders: "table-cell-document",
  _resolvedBorders: "table-cell-document",
  margins: "table-cell-document",
  _resolvedMargins: "table-cell-document",
  _originalFormatting: "table-cell-document",
  tcPrChange: "semantic",
  cellMarker: "semantic",
  _preserveVMergeRestart: "table-cell-document",
  _docxVMergeContinuationCells: "semantic",
} as const satisfies Readonly<Record<keyof TableCellAttrs, SourceIdentityAttrDisposition>>);

/**
 * Text-box group and anchor ids are random per editor load. Their equality
 * relationships carry source structure; their literal values do not. Keeping
 * this map total makes every new text-box attribute an explicit identity
 * decision instead of letting projection drift become invisible.
 */
const TEXT_BOX_SOURCE_IDENTITY_ATTRS = Object.freeze({
  width: "semantic",
  height: "semantic",
  autoFit: "semantic",
  textWrap: "semantic",
  textBoxId: "semantic",
  fillColor: "semantic",
  outlineWidth: "semantic",
  outlineColor: "semantic",
  outlineStyle: "semantic",
  transform: "semantic",
  marginTop: "semantic",
  marginBottom: "semantic",
  marginLeft: "semantic",
  marginRight: "semantic",
  verticalAlign: "semantic",
  displayMode: "semantic",
  cssFloat: "semantic",
  wrapType: "semantic",
  wrapText: "semantic",
  distTop: "semantic",
  distBottom: "semantic",
  distLeft: "semantic",
  distRight: "semantic",
  position: "semantic",
  _docxPlacement: "semantic",
  _docxGroupId: "text-box-group",
  _docxAnchorId: "text-box-anchor",
  _docxTrackedChange: "semantic",
  _docxInlineSdts: "semantic",
} as const satisfies Readonly<Record<keyof TextBoxAttrs, SourceIdentityAttrDisposition>>);

const TEXT_BOX_ANCHOR_SOURCE_IDENTITY_ATTRS = Object.freeze({
  anchorId: "text-box-anchor",
} as const satisfies Readonly<Record<keyof TextBoxAnchorAttrs, SourceIdentityAttrDisposition>>);

type AlphaIdentityState = {
  readonly left: Map<string, number>;
  readonly right: Map<string, number>;
};

type SourceIdentityState = {
  readonly anchors: AlphaIdentityState;
  readonly groups: AlphaIdentityState;
};

const SOURCE_IDENTITY_ATTRS_BY_NODE = Object.freeze({
  paragraph: PARAGRAPH_SOURCE_IDENTITY_ATTRS,
  table: TABLE_SOURCE_IDENTITY_ATTRS,
  tableRow: TABLE_ROW_SOURCE_IDENTITY_ATTRS,
  tableCell: TABLE_CELL_SOURCE_IDENTITY_ATTRS,
  tableHeader: TABLE_CELL_SOURCE_IDENTITY_ATTRS,
  textBox: TEXT_BOX_SOURCE_IDENTITY_ATTRS,
  textBoxAnchor: TEXT_BOX_ANCHOR_SOURCE_IDENTITY_ATTRS,
});

type GovernedSourceIdentityNodeName = keyof typeof SOURCE_IDENTITY_ATTRS_BY_NODE;

const governedNodeName = (value: string): GovernedSourceIdentityNodeName | null => {
  switch (value) {
    case "paragraph":
    case "table":
    case "tableRow":
    case "tableCell":
    case "tableHeader":
    case "textBox":
    case "textBoxAnchor":
      return value;
    default:
      return null;
  }
};

const sourceIdentityAttrDisposition = (
  value: unknown,
  nodeTypeName: GovernedSourceIdentityNodeName,
  key: string,
): SourceIdentityAttrDisposition => {
  switch (value) {
    case "semantic":
    case "transport":
    case "presentation":
    case "paragraph-document":
    case "paragraph-section":
    case "paragraph-bookmark-boundaries":
    case "table-document":
    case "table-row-document":
    case "table-cell-document":
    case "text-box-anchor":
    case "text-box-group":
      return value;
    default:
      return panic("A governed ProseMirror source attribute has no identity disposition", {
        nodeTypeName,
        key,
      });
  }
};

/** @internal Exposed for the schema-policy completeness invariant. */
export const sourceIdentityGovernedNodeNames = (): readonly string[] =>
  Object.keys(SOURCE_IDENTITY_ATTRS_BY_NODE);

/** @internal Exposed for the schema-policy completeness invariant. */
export const sourceIdentityGovernedAttrKeys = (nodeTypeName: string): readonly string[] => {
  const name = governedNodeName(nodeTypeName);
  return name === null ? [] : Object.keys(SOURCE_IDENTITY_ATTRS_BY_NODE[name]);
};

/**
 * Every governed schema attribute requires an explicit disposition. Unknown
 * node kinds remain exact; an unknown attr on a governed kind is a schema bug.
 */
export const sourceIdentityAttrDispositionFor = (
  nodeTypeName: string,
  key: string,
): SourceIdentityAttrDisposition => {
  const name = governedNodeName(nodeTypeName);
  if (name === null) return "semantic";
  return sourceIdentityAttrDisposition(
    Reflect.get(SOURCE_IDENTITY_ATTRS_BY_NODE[name], key),
    name,
    key,
  );
};

const alphaIdentity = (value: unknown, identities: Map<string, number>): unknown => {
  if (value === null || value === undefined) return null;
  if (typeof value !== "string") return value;
  const existing = identities.get(value);
  if (existing !== undefined) return existing;
  const identity = identities.size;
  identities.set(value, identity);
  return identity;
};

const validAlphaIdentity = (value: unknown, optional: boolean): boolean =>
  (optional && (value === null || value === undefined)) ||
  (typeof value === "string" && value.length > 0);

const paragraphDocumentProjection = (node: PMNode): unknown => {
  const attrs = expectParagraphAttrs(node);
  return {
    formatting: paragraphAttrsToFormatting(attrs) ?? null,
    listRendering: listRenderingFromAttrs(attrs) ?? null,
  };
};

const paragraphSectionProjection = (node: PMNode): unknown => {
  const attrs = expectParagraphAttrs(node);
  if (attrs._sectionProperties !== undefined && attrs._sectionProperties !== null) {
    return attrs._sectionProperties;
  }
  return attrs.sectionBreakType ? { sectionStart: attrs.sectionBreakType } : null;
};

const documentProjection = (node: PMNode, disposition: SourceIdentityAttrDisposition): unknown => {
  switch (disposition) {
    case "paragraph-document":
      return paragraphDocumentProjection(node);
    case "paragraph-section":
      return paragraphSectionProjection(node);
    case "table-document":
      return tableAttrsToFormatting(expectTableAttrs(node)) ?? null;
    case "table-row-document":
      return tableRowAttrsToFormatting(expectTableRowAttrs(node)) ?? null;
    case "table-cell-document":
      return tableCellAttrsToFormatting(expectTableCellAttrs(node)) ?? null;
    default:
      return panic("A source identity disposition has no Document projection", {
        disposition,
      });
  }
};

const attrsMatch = (left: PMNode, right: PMNode, state: SourceIdentityState): string | null => {
  const attrKeys = new Set([...Object.keys(left.attrs), ...Object.keys(right.attrs)]);
  let comparedDocumentProjections: Set<SourceIdentityAttrDisposition> | null = null;
  const governedName = governedNodeName(left.type.name);
  for (const key of attrKeys) {
    const disposition =
      governedName === null
        ? "semantic"
        : sourceIdentityAttrDisposition(
            Reflect.get(SOURCE_IDENTITY_ATTRS_BY_NODE[governedName], key),
            governedName,
            key,
          );
    const leftValue = left.attrs[key];
    const rightValue = right.attrs[key];
    if (disposition === "presentation" || disposition === "paragraph-bookmark-boundaries") {
      continue;
    }
    const rawValuesMatch = canonicalJson(leftValue) === canonicalJson(rightValue);
    if (
      disposition === "paragraph-document" ||
      disposition === "paragraph-section" ||
      disposition === "table-document" ||
      disposition === "table-row-document" ||
      disposition === "table-cell-document"
    ) {
      if (rawValuesMatch) continue;
      if (comparedDocumentProjections?.has(disposition)) continue;
      comparedDocumentProjections ??= new Set();
      comparedDocumentProjections.add(disposition);
      if (
        canonicalJson(documentProjection(left, disposition)) !==
        canonicalJson(documentProjection(right, disposition))
      ) {
        return `$${disposition}`;
      }
      continue;
    }
    if (disposition === "semantic" || disposition === "transport") {
      if (!rawValuesMatch) return key;
      continue;
    }
    const identityIsOptional = left.type.name === "textBox";
    if (
      !validAlphaIdentity(leftValue, identityIsOptional) ||
      !validAlphaIdentity(rightValue, identityIsOptional)
    ) {
      return key;
    }
    const identities = disposition === "text-box-group" ? state.groups : state.anchors;
    if (alphaIdentity(leftValue, identities.left) !== alphaIdentity(rightValue, identities.right)) {
      return key;
    }
  }
  return null;
};

const sourceIdentityBookmarkChildren = (node: PMNode): readonly PMNode[] | null => {
  if (node.type.name !== "paragraph") return null;
  const bookmarks = expectParagraphAttrs(node).bookmarks;
  if (!bookmarks || bookmarks.length === 0) return null;
  const boundaryType = node.type.schema.nodes["bookmarkBoundary"];
  if (!boundaryType) {
    return panic("The paragraph schema has no bookmark-boundary projection");
  }
  const children: PMNode[] = [];
  node.forEach((child) => {
    children.push(child);
  });
  return [
    ...bookmarks.map(({ id, name }) => boundaryType.create({ type: "start", id, name })),
    ...children,
    ...bookmarks.map(({ id }) => boundaryType.create({ type: "end", id })),
  ];
};

const firstDifferencePath = (
  left: PMNode,
  right: PMNode,
  state: SourceIdentityState,
  path: string,
): string => {
  if (left.type !== right.type) return `${path}.type`;
  const attrDifference = attrsMatch(left, right, state);
  if (attrDifference !== null) return `${path}.attrs.${attrDifference}`;
  if (left.text !== right.text) return `${path}.text`;
  if (
    canonicalJson(left.marks.map((mark) => mark.toJSON())) !==
    canonicalJson(right.marks.map((mark) => mark.toJSON()))
  ) {
    return `${path}.marks`;
  }
  const leftBookmarkChildren = sourceIdentityBookmarkChildren(left);
  const rightBookmarkChildren = sourceIdentityBookmarkChildren(right);
  const leftChildCount = leftBookmarkChildren?.length ?? left.childCount;
  const rightChildCount = rightBookmarkChildren?.length ?? right.childCount;
  if (leftChildCount !== rightChildCount) return `${path}.childCount`;
  for (let index = 0; index < leftChildCount; index++) {
    const leftChild = leftBookmarkChildren?.at(index) ?? left.child(index);
    const rightChild = rightBookmarkChildren?.at(index) ?? right.child(index);
    const difference = firstDifferencePath(
      leftChild,
      rightChild,
      state,
      `${path}.content[${String(index)}]`,
    );
    if (difference !== "") return difference;
  }
  return "";
};

/**
 * Return the first structural or semantic PM difference. Session-random text
 * box identifiers are alpha-compared, so grouping and anchor links must remain
 * identical even though independently loaded documents have different salts.
 */
export const firstProseMirrorSourceIdentityDifferencePath = (left: PMNode, right: PMNode): string =>
  firstDifferencePath(
    left,
    right,
    {
      anchors: { left: new Map(), right: new Map() },
      groups: { left: new Map(), right: new Map() },
    },
    "doc",
  );
