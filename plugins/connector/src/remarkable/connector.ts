import type { Entry, RemarkableApi, SimpleEntry } from "rmapi-js";

import {
  FOLIO_CONNECTOR_ROOT_ID,
  type FolioConnectorDocumentFormat,
  type FolioConnectorItem,
  type FolioConnectorItemKind,
  type FolioConnectorParentId,
  type FolioConnectorUploadMimeType,
  type FolioCloudConnector,
} from "../types";
import {
  RemarkableConnectorDownloadError,
  RemarkableConnectorUploadError,
} from "../errors";

export const REMARKABLE_CONNECTOR_PROVIDER_ID = "remarkable" as const;

export type RemarkableConnector = FolioCloudConnector & {
  readonly providerId: typeof REMARKABLE_CONNECTOR_PROVIDER_ID;
  /** Escape hatch to the underlying rmapi-js client. */
  readonly api: RemarkableApi;
};

const toParentId = (parent: string | undefined): FolioConnectorParentId =>
  parent === undefined || parent === "" ? FOLIO_CONNECTOR_ROOT_ID : parent;

const mapEntryKind = (entry: Entry): FolioConnectorItemKind => {
  if (entry.type === "CollectionType") {
    return "folder";
  }
  if (entry.type === "TemplateType") {
    return "template";
  }
  return "document";
};

export const mapRemarkableEntry = (entry: Entry): FolioConnectorItem => {
  const kind = mapEntryKind(entry);
  const documentFormat =
    entry.type === "DocumentType" ? entry.fileType : undefined;

  return {
    id: entry.id,
    hash: entry.hash,
    name: entry.visibleName,
    parentId: toParentId(entry.parent),
    kind,
    modifiedAt: entry.lastModified,
    starred: entry.pinned,
    ...(documentFormat ? { documentFormat } : {}),
  };
};

const mapSimpleEntry = (
  entry: SimpleEntry,
  fallback: Pick<FolioConnectorItem, "name" | "parentId" | "kind" | "documentFormat">,
): FolioConnectorItem => ({
  id: entry.id,
  hash: entry.hash,
  name: fallback.name,
  parentId: fallback.parentId,
  kind: fallback.kind,
  modifiedAt: new Date().toISOString(),
  starred: false,
  ...(fallback.documentFormat ? { documentFormat: fallback.documentFormat } : {}),
});

const refreshItemHash = (
  item: FolioConnectorItem,
  nextHash: string,
): FolioConnectorItem => ({
  ...item,
  hash: nextHash,
});

const uploadMimeToDocumentFormat = (
  mimeType: FolioConnectorUploadMimeType,
): FolioConnectorDocumentFormat => {
  if (mimeType === "application/pdf") {
    return "pdf";
  }
  return "epub";
};

export const createRemarkableConnector = (api: RemarkableApi): RemarkableConnector => ({
  providerId: REMARKABLE_CONNECTOR_PROVIDER_ID,
  api,

  async listItems(options = {}) {
    const entries = await api.listItems(options.refresh);
    const mapped = entries.map(mapRemarkableEntry);
    if (options.parentId === undefined) {
      return mapped;
    }
    return mapped.filter((item) => item.parentId === options.parentId);
  },

  async uploadDocument({ name, bytes, mimeType, parentId = FOLIO_CONNECTOR_ROOT_ID }) {
    const documentFormat = uploadMimeToDocumentFormat(mimeType);
    let entry;
    if (mimeType === "application/pdf") {
      entry = await api.uploadPdf(name, bytes);
    } else if (mimeType === "application/epub+zip") {
      entry = await api.uploadEpub(name, bytes);
    } else {
      throw new RemarkableConnectorUploadError({
        message: "Unsupported upload mime type",
        mimeType,
      });
    }

    let hash = entry.hash;
    if (parentId !== FOLIO_CONNECTOR_ROOT_ID) {
      const moved = await api.move(entry.hash, parentId);
      hash = moved.hash;
    }

    return mapSimpleEntry(
      { id: entry.id, hash },
      {
        name,
        parentId,
        kind: "document",
        documentFormat,
      },
    );
  },

  async downloadDocument(item) {
    if (item.kind !== "document") {
      throw new RemarkableConnectorDownloadError({
        message: "Only documents can be downloaded",
        itemId: item.id,
      });
    }

    if (item.documentFormat === "pdf") {
      return await api.getPdf(item.id, item.hash);
    }
    if (item.documentFormat === "epub") {
      return await api.getEpub(item.id, item.hash);
    }

    throw new RemarkableConnectorDownloadError({
      message:
        "Notebook downloads are not supported yet; export the notebook from reMarkable as PDF first",
      itemId: item.id,
      ...(item.documentFormat
        ? { documentFormat: item.documentFormat }
        : {}),
    });
  },

  async createFolder({ name, parentId = FOLIO_CONNECTOR_ROOT_ID }) {
    const entry =
      parentId === FOLIO_CONNECTOR_ROOT_ID
        ? await api.uploadFolder(name)
        : await api.putFolder(name, { parent: parentId });

    return mapSimpleEntry(entry, {
      name,
      parentId,
      kind: "folder",
    });
  },

  async moveItem(item, parentId) {
    const moved = await api.move(item.hash, parentId);
    return refreshItemHash({ ...item, parentId }, moved.hash);
  },

  async renameItem(item, name) {
    const renamed = await api.rename(item.hash, name);
    return refreshItemHash({ ...item, name }, renamed.hash);
  },

  async deleteItem(item) {
    await api.delete(item.hash);
  },

  async setStarred(item, starred) {
    const updated = await api.stared(item.hash, starred);
    return refreshItemHash({ ...item, starred }, updated.hash);
  },
});
