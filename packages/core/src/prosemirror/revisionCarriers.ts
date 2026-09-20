import type { Node as PMNode } from "prosemirror-model";

import {
  PARAGRAPH_MARK_CHANGE_KINDS,
  type ParagraphMarkChangeKind,
  type PropertyRevisionKind,
} from "@stll/docx-core/model";

/**
 * Where each tracked property revision sits in the editable document, and how
 * it resolves.
 *
 * One table, total over `PROPERTY_REVISION_KINDS`, because the alternative is
 * the same list of element names written out again in the carrier reader, the
 * accept/reject command, the tracked-change list and the comparison — four
 * places that drifted apart the moment a fifth revision was modelled. A
 * revision the model gains is a decision here and a compile error at every
 * consumer of the derived unions below.
 */
type PropertyRevisionSite =
  | {
      /**
       * Carried on a run's `runPropertyChange` mark rather than on a node, so
       * the inline resolver owns reading and resolving it.
       */
      resolution: "inline-mark";
      kind: PropertyRevisionKind;
    }
  | {
      /**
       * Carried on a node attr and resolved by the shared node-attr pass:
       * accept drops the matched records and keeps the live properties, reject
       * additionally applies the stored previous set wholesale.
       */
      resolution: "node-attrs";
      kind: PropertyRevisionKind;
      /** PM node types whose attrs carry the records. */
      nodeTypes: readonly string[];
      /** The attr holding the record array, read and written by name. */
      attr: string;
      /** The name the reviewer, the AI-edit reader and the sidebars know it by. */
      carrier: string;
      range: "whole-node";
    }
  | {
      /**
       * Carried on a paragraph attr and resolved by the paragraph's own pass,
       * which rebases style-derived attrs and restores header/footer
       * references the generic patch cannot reach.
       */
      resolution: "paragraph-pass";
      kind: PropertyRevisionKind;
      nodeTypes: readonly ["paragraph"];
      /** Path from the paragraph's attrs down to the record array. */
      path: readonly [string, ...string[]];
      carrier: string;
      /**
       * A paragraph's property revisions resolve over its mark boundary, the
       * range a `w:pPrChange` card selects, not over the whole block.
       */
      range: "paragraph-boundary";
    };

export const PROPERTY_REVISION_SITES = {
  runPropertyChange: { kind: "runPropertyChange", resolution: "inline-mark" },
  paragraphPropertyChange: {
    kind: "paragraphPropertyChange",
    resolution: "paragraph-pass",
    nodeTypes: ["paragraph"],
    path: ["_propertyChanges"],
    carrier: "paragraphPropertiesChanged",
    range: "paragraph-boundary",
  },
  sectionPropertyChange: {
    kind: "sectionPropertyChange",
    resolution: "paragraph-pass",
    nodeTypes: ["paragraph"],
    path: ["_sectionProperties", "propertyChanges"],
    carrier: "sectionPropertiesChanged",
    range: "paragraph-boundary",
  },
  tablePropertyChange: {
    kind: "tablePropertyChange",
    resolution: "node-attrs",
    nodeTypes: ["table"],
    attr: "tblPrChange",
    carrier: "tablePropertiesChanged",
    range: "whole-node",
  },
  tablePropertyExceptionChange: {
    kind: "tablePropertyExceptionChange",
    resolution: "node-attrs",
    nodeTypes: ["tableRow"],
    attr: "tblPrExChange",
    carrier: "tablePropertyExceptionsChanged",
    range: "whole-node",
  },
  tableRowPropertyChange: {
    kind: "tableRowPropertyChange",
    resolution: "node-attrs",
    nodeTypes: ["tableRow"],
    attr: "trPrChange",
    carrier: "rowPropertiesChanged",
    range: "whole-node",
  },
  tableCellPropertyChange: {
    kind: "tableCellPropertyChange",
    resolution: "node-attrs",
    nodeTypes: ["tableCell", "tableHeader"],
    attr: "tcPrChange",
    carrier: "cellPropertiesChanged",
    range: "whole-node",
  },
  // The mapped `satisfies` binds each entry's key to its own `kind`, so a site
  // cannot name a revision other than the one it is filed under, and the kind
  // survives into the derived unions below.
} as const satisfies { [Kind in PropertyRevisionKind]: PropertyRevisionSite & { kind: Kind } };

type SiteOf<TKind extends PropertyRevisionKind> = (typeof PROPERTY_REVISION_SITES)[TKind];

/** One property revision carried on a node's attrs rather than on a mark. */
export type NodePropertyRevisionSite = Extract<SiteOf<PropertyRevisionKind>, { carrier: string }>;

/** The sites the shared node-attr pass resolves, and therefore must patch. */
export type NodeAttrsPropertyRevisionSite = Extract<
  NodePropertyRevisionSite,
  { resolution: "node-attrs" }
>;

/** The kinds that pass resolves. */
export type NodeAttrsPropertyRevisionKind = NodeAttrsPropertyRevisionSite["kind"];

/** The reviewer-facing name of a property revision carried on a node. */
export type PropertyRevisionCarrier = NodePropertyRevisionSite["carrier"];

export type FolioNodeRevisionKind =
  | "paragraphMarkInserted"
  | "paragraphMarkDeleted"
  | PropertyRevisionCarrier;

const NODE_PROPERTY_REVISION_SITES = Object.values(PROPERTY_REVISION_SITES).flatMap((site) =>
  site.resolution === "inline-mark" ? [] : [site],
);

/** The property revisions a PM node of this type can carry, in element order. */
export const nodePropertyRevisionSites = (nodeTypeName: string): NodePropertyRevisionSite[] =>
  NODE_PROPERTY_REVISION_SITES.filter(({ nodeTypes }) =>
    nodeTypes.some((name) => name === nodeTypeName),
  );

/**
 * Which carrier each paragraph-mark kind reports as. A relocation's break
 * resolves as the insertion or deletion it is; the kind exists so a reader is
 * told the two ends belong together, not to change what resolving does.
 *
 * Total over the kinds by construction, so a kind the model gains has to be
 * given a carrier here rather than silently reporting none.
 */
const PARAGRAPH_MARK_KIND_CARRIERS = {
  ins: "paragraphMarkInserted",
  moveTo: "paragraphMarkInserted",
  del: "paragraphMarkDeleted",
  moveFrom: "paragraphMarkDeleted",
} as const satisfies Record<ParagraphMarkChangeKind, FolioNodeRevisionKind>;

const isParagraphMarkChangeKind = (value: unknown): value is ParagraphMarkChangeKind =>
  PARAGRAPH_MARK_CHANGE_KINDS.some((kind) => kind === value);

export type FolioNodeRevisionCarrier = {
  id: number;
  type: FolioNodeRevisionKind;
  author: string;
  date: string | null;
  text: string;
  from: number;
  to: number;
};

type RevisionMetadata = Pick<FolioNodeRevisionCarrier, "id" | "author" | "date">;

const isObjectRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null;

/** The `w:id`, author and date a stored revision's `info` carries, or null. */
export const propertyRevisionMetadata = (value: unknown): RevisionMetadata | null => {
  if (!isObjectRecord(value) || typeof value["id"] !== "number") {
    return null;
  }
  return {
    id: value["id"],
    author: typeof value["author"] === "string" ? value["author"] : "",
    date: typeof value["date"] === "string" ? value["date"] : null,
  };
};

/**
 * A stored property-revision record as ProseMirror attrs carry it. The
 * previous property set is whatever the site's own element stores, which the
 * reject patch filed under that kind narrows.
 */
export type StoredPropertyRevision = {
  info?: unknown;
  previousFormatting?: unknown;
};

/** The record array a site names, walked from a node's attrs. */
export const propertyRevisionRecords = (
  node: PMNode,
  site: NodePropertyRevisionSite,
): StoredPropertyRevision[] => {
  const path = site.resolution === "node-attrs" ? [site.attr] : site.path;
  let value: unknown = node.attrs;
  for (const key of path) {
    if (!isObjectRecord(value)) {
      return [];
    }
    value = value[key];
  }
  return Array.isArray(value) ? value.filter(isObjectRecord) : [];
};

type AppendPropertyCarriersOptions = {
  carriers: FolioNodeRevisionCarrier[];
  changes: readonly unknown[];
  type: FolioNodeRevisionKind;
  node: PMNode;
  from: number;
  to: number;
};

const appendPropertyCarriers = ({
  carriers,
  changes,
  type,
  node,
  from,
  to,
}: AppendPropertyCarriersOptions): void => {
  let text: string | null = null;
  for (const change of changes) {
    if (!isObjectRecord(change)) {
      continue;
    }
    const metadata = propertyRevisionMetadata(change["info"]);
    if (!metadata) {
      continue;
    }
    text ??= node.textContent;
    carriers.push({ ...metadata, type, text, from, to });
  }
};

/**
 * Read revision records stored on ProseMirror node attributes and return the
 * exact range the shared accept/reject command requires for each carrier.
 */
export const getFolioNodeRevisionCarriers = (
  node: PMNode,
  nodePos: number,
): FolioNodeRevisionCarrier[] => {
  const carriers: FolioNodeRevisionCarrier[] = [];

  if (node.type.name === "paragraph") {
    const paragraphMark = node.attrs["pPrMark"];
    const paragraphMarkKind =
      isObjectRecord(paragraphMark) && isParagraphMarkChangeKind(paragraphMark["kind"])
        ? PARAGRAPH_MARK_KIND_CARRIERS[paragraphMark["kind"]]
        : undefined;
    if (paragraphMarkKind !== undefined && isObjectRecord(paragraphMark)) {
      const metadata = propertyRevisionMetadata(paragraphMark["info"]);
      if (metadata) {
        carriers.push({
          ...metadata,
          type: paragraphMarkKind,
          text: node.textContent,
          from: nodePos + node.nodeSize - 1,
          to: nodePos + node.nodeSize,
        });
      }
    }
  }

  for (const site of nodePropertyRevisionSites(node.type.name)) {
    const range =
      site.range === "paragraph-boundary"
        ? { from: nodePos + node.nodeSize - 1, to: nodePos + node.nodeSize }
        : { from: nodePos, to: nodePos + node.nodeSize };
    appendPropertyCarriers({
      carriers,
      changes: propertyRevisionRecords(node, site),
      type: site.carrier,
      node,
      ...range,
    });
  }

  return carriers;
};
