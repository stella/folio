import type { Node as PMNode } from "prosemirror-model";

import type {
  TableCellFormatting,
  TableFormatting,
  TablePropertyExceptionFormatting,
  TableRowFormatting,
} from "../../types/document";
import {
  nodePropertyRevisionSites,
  propertyRevisionMetadata,
  propertyRevisionRecords,
  type NodeAttrsPropertyRevisionKind,
  type NodeAttrsPropertyRevisionSite,
} from "../revisionCarriers";
import {
  tableCellRejectAttrPatch,
  tablePropertyExceptionsRejectAttrPatch,
  tableRejectAttrPatch,
  tableRowRejectAttrPatch,
} from "./propertyChangeScope";

type NodeAttrsRejectPatch = (previousFormatting: unknown, node: PMNode) => Record<string, unknown>;

/** The old property set each node-attr revision restores on rejection. */
const NODE_ATTRS_REJECT_PATCHES = {
  tablePropertyChange: (previousFormatting) =>
    tableRejectAttrPatch(previousFormatting as TableFormatting | undefined),
  tablePropertyExceptionChange: (previousFormatting) =>
    tablePropertyExceptionsRejectAttrPatch(
      previousFormatting as TablePropertyExceptionFormatting | undefined,
    ),
  tableRowPropertyChange: (previousFormatting) =>
    tableRowRejectAttrPatch(previousFormatting as TableRowFormatting | undefined),
  tableCellPropertyChange: (previousFormatting, node) =>
    tableCellRejectAttrPatch(
      previousFormatting as TableCellFormatting | undefined,
      node.attrs["_originalFormatting"] as TableCellFormatting | null | undefined,
    ),
} as const satisfies Record<NodeAttrsPropertyRevisionKind, NodeAttrsRejectPatch>;

/** Resolve one property-change site without a transaction. */
type ResolveNodePropertyChangeAttrsOptions = {
  node: PMNode;
  site: NodeAttrsPropertyRevisionSite;
  mode: "accept" | "reject";
  revisionSet: ReadonlySet<number> | null;
};
export const resolveNodePropertyChangeAttrs = ({
  node,
  site,
  mode,
  revisionSet,
}: ResolveNodePropertyChangeAttrsOptions): Record<string, unknown> | null => {
  const changes = propertyRevisionRecords(node, site);
  if (changes.length === 0) {
    return null;
  }
  const matches = changes.filter((change) => {
    const metadata = propertyRevisionMetadata(change.info);
    return revisionSet === null || (metadata !== null && revisionSet.has(metadata.id));
  });
  if (matches.length === 0) {
    return null;
  }
  const remaining =
    revisionSet === null ? [] : changes.filter((change) => !matches.includes(change));
  const nextAttrs: Record<string, unknown> = {
    ...node.attrs,
    [site.attr]: remaining.length > 0 ? remaining : null,
  };
  if (mode === "reject") {
    const rejectPatch = NODE_ATTRS_REJECT_PATCHES[site.kind];
    for (const change of matches.toReversed()) {
      Object.assign(nextAttrs, rejectPatch(change.previousFormatting, node));
    }
  }
  return nextAttrs;
};

/** Resolve every node-attr site, including both independent sites on a row. */
export const resolveAllNodePropertyChangeAttrs = (
  node: PMNode,
  mode: "accept" | "reject",
): PMNode => {
  let current = node;
  for (const site of nodePropertyRevisionSites(node.type.name)) {
    if (site.resolution !== "node-attrs") {
      continue;
    }
    const attrs = resolveNodePropertyChangeAttrs({ node: current, site, mode, revisionSet: null });
    if (attrs !== null) {
      current = current.type.create(attrs, current.content, current.marks);
    }
  }
  return current;
};
