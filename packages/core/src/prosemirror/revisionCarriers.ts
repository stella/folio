import type { Node as PMNode } from "prosemirror-model";

import { PARAGRAPH_MARK_CHANGE_KINDS, type ParagraphMarkChangeKind } from "@stll/docx-core/model";

import { expectParagraphAttrs } from "./attrs";

export type FolioNodeRevisionKind =
  | "paragraphMarkInserted"
  | "paragraphMarkDeleted"
  | "paragraphPropertiesChanged"
  | "sectionPropertiesChanged"
  | "tablePropertiesChanged"
  | "rowPropertiesChanged"
  | "cellPropertiesChanged";

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

const revisionMetadata = (value: unknown): RevisionMetadata | null => {
  if (!isObjectRecord(value) || typeof value["id"] !== "number") {
    return null;
  }
  return {
    id: value["id"],
    author: typeof value["author"] === "string" ? value["author"] : "",
    date: typeof value["date"] === "string" ? value["date"] : null,
  };
};

type AppendPropertyCarriersOptions = {
  carriers: FolioNodeRevisionCarrier[];
  changes: unknown;
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
  if (!Array.isArray(changes)) {
    return;
  }
  let text: string | null = null;
  for (const change of changes) {
    if (!isObjectRecord(change)) {
      continue;
    }
    const metadata = revisionMetadata(change["info"]);
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
    const from = nodePos + node.nodeSize - 1;
    const to = nodePos + node.nodeSize;
    const paragraphAttrs = expectParagraphAttrs(node);
    const paragraphMark = node.attrs["pPrMark"];
    const paragraphMarkKind =
      isObjectRecord(paragraphMark) && isParagraphMarkChangeKind(paragraphMark["kind"])
        ? PARAGRAPH_MARK_KIND_CARRIERS[paragraphMark["kind"]]
        : undefined;
    if (paragraphMarkKind !== undefined && isObjectRecord(paragraphMark)) {
      const metadata = revisionMetadata(paragraphMark["info"]);
      if (metadata) {
        carriers.push({
          ...metadata,
          type: paragraphMarkKind,
          text: node.textContent,
          from,
          to,
        });
      }
    }
    appendPropertyCarriers({
      carriers,
      changes: paragraphAttrs._propertyChanges,
      type: "paragraphPropertiesChanged",
      node,
      from,
      to,
    });
    const sectionProperties = node.attrs["_sectionProperties"];
    appendPropertyCarriers({
      carriers,
      changes: isObjectRecord(sectionProperties) ? sectionProperties["propertyChanges"] : null,
      type: "sectionPropertiesChanged",
      node,
      from,
      to,
    });
    return carriers;
  }

  const range = { from: nodePos, to: nodePos + node.nodeSize };
  switch (node.type.name) {
    case "table":
      appendPropertyCarriers({
        carriers,
        changes: node.attrs["tblPrChange"],
        type: "tablePropertiesChanged",
        node,
        ...range,
      });
      break;
    case "tableRow":
      appendPropertyCarriers({
        carriers,
        changes: node.attrs["trPrChange"],
        type: "rowPropertiesChanged",
        node,
        ...range,
      });
      break;
    case "tableCell":
    case "tableHeader":
      appendPropertyCarriers({
        carriers,
        changes: node.attrs["tcPrChange"],
        type: "cellPropertiesChanged",
        node,
        ...range,
      });
      break;
    default:
      break;
  }
  return carriers;
};
