/** Root collection id in reMarkable's cloud file tree. */
export const FOLIO_CONNECTOR_ROOT_ID = "" as const;

/** Trash collection id in reMarkable's cloud file tree. */
export const FOLIO_CONNECTOR_TRASH_ID = "trash" as const;

export type FolioConnectorParentId =
  | typeof FOLIO_CONNECTOR_ROOT_ID
  | typeof FOLIO_CONNECTOR_TRASH_ID
  | (string & {});

export type FolioConnectorItemKind = "folder" | "document" | "template";

export type FolioConnectorDocumentFormat = "pdf" | "epub" | "notebook";

export type FolioConnectorUploadMimeType = "application/pdf" | "application/epub+zip";

/**
 * Provider-neutral view of a remote item. Every connector maps its native entry
 * shape into this so hosts can drive multiple backends with one UI.
 */
export type FolioConnectorItem = {
  id: string;
  hash: string;
  name: string;
  parentId: FolioConnectorParentId;
  kind: FolioConnectorItemKind;
  modifiedAt: string;
  starred: boolean;
  documentFormat?: FolioConnectorDocumentFormat;
};

export type FolioConnectorListOptions = {
  /** When set, only children of this folder are returned. */
  parentId?: FolioConnectorParentId;
  /** Ask the provider to refresh its remote index before listing. */
  refresh?: boolean;
};

export type FolioConnectorUploadOptions = {
  name: string;
  bytes: Uint8Array;
  mimeType: FolioConnectorUploadMimeType;
  parentId?: FolioConnectorParentId;
};

export type FolioConnectorCreateFolderOptions = {
  name: string;
  parentId?: FolioConnectorParentId;
};

/**
 * The structural contract every folio cloud connector implements. Host apps
 * (Stella, scripts, agents) call these methods to sync documents without
 * depending on a provider SDK.
 */
export type FolioCloudConnector = {
  readonly providerId: string;
  listItems(options?: FolioConnectorListOptions): Promise<FolioConnectorItem[]>;
  uploadDocument(options: FolioConnectorUploadOptions): Promise<FolioConnectorItem>;
  downloadDocument(item: FolioConnectorItem): Promise<Uint8Array>;
  createFolder(options: FolioConnectorCreateFolderOptions): Promise<FolioConnectorItem>;
  moveItem(item: FolioConnectorItem, parentId: FolioConnectorParentId): Promise<FolioConnectorItem>;
  renameItem(item: FolioConnectorItem, name: string): Promise<FolioConnectorItem>;
  deleteItem(item: FolioConnectorItem): Promise<void>;
  setStarred(item: FolioConnectorItem, starred: boolean): Promise<FolioConnectorItem>;
};
