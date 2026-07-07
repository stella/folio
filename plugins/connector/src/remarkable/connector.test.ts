import { describe, expect, test } from "bun:test";
import type { Entry, RemarkableApi } from "rmapi-js";

import { createRemarkableConnector, mapRemarkableEntry } from "./connector";
import { FOLIO_CONNECTOR_ROOT_ID } from "../types";

const samplePdf: Entry = {
  id: "doc-1",
  hash: "hash-doc-1",
  visibleName: "Contract.pdf",
  lastModified: "2026-07-01T12:00:00.000Z",
  pinned: false,
  parent: "",
  type: "DocumentType",
  fileType: "pdf",
  lastOpened: "2026-07-01T12:00:00.000Z",
};

const sampleFolder: Entry = {
  id: "folder-1",
  hash: "hash-folder-1",
  visibleName: "Matters",
  lastModified: "2026-06-30T08:00:00.000Z",
  pinned: true,
  parent: "",
  type: "CollectionType",
};

const createMockApi = (overrides: Partial<RemarkableApi> = {}): RemarkableApi =>
  ({
    listItems: async () => [sampleFolder, samplePdf],
    uploadPdf: async () => ({ id: "doc-2", hash: "hash-doc-2" }),
    uploadEpub: async () => ({ id: "doc-3", hash: "hash-doc-3" }),
    uploadFolder: async () => ({ id: "folder-2", hash: "hash-folder-2" }),
    putFolder: async () => ({ id: "folder-3", hash: "hash-folder-3" }),
    move: async () => ({ hash: "hash-moved" }),
    rename: async () => ({ hash: "hash-renamed" }),
    delete: async () => ({ hash: "hash-deleted" }),
    stared: async () => ({ hash: "hash-starred" }),
    getPdf: async () => new Uint8Array([0x25, 0x50, 0x44, 0x46]),
    getEpub: async () => new Uint8Array([0x50, 0x4b, 0x03, 0x04]),
    ...overrides,
  }) as RemarkableApi;

describe("mapRemarkableEntry", () => {
  test("maps folders and documents into connector items", () => {
    expect(mapRemarkableEntry(sampleFolder)).toEqual({
      id: "folder-1",
      hash: "hash-folder-1",
      name: "Matters",
      parentId: FOLIO_CONNECTOR_ROOT_ID,
      kind: "folder",
      modifiedAt: "2026-06-30T08:00:00.000Z",
      starred: true,
    });

    expect(mapRemarkableEntry(samplePdf)).toEqual({
      id: "doc-1",
      hash: "hash-doc-1",
      name: "Contract.pdf",
      parentId: FOLIO_CONNECTOR_ROOT_ID,
      kind: "document",
      modifiedAt: "2026-07-01T12:00:00.000Z",
      starred: false,
      documentFormat: "pdf",
    });
  });
});

describe("createRemarkableConnector", () => {
  test("filters listItems by parentId", async () => {
    const nestedPdf: Entry = {
      ...samplePdf,
      id: "doc-nested",
      hash: "hash-doc-nested",
      parent: "folder-1",
    };
    const connector = createRemarkableConnector(
      createMockApi({
        listItems: async () => [sampleFolder, nestedPdf],
      }),
    );

    const rootItems = await connector.listItems();
    expect(rootItems).toHaveLength(2);

    const folderChildren = await connector.listItems({ parentId: "folder-1" });
    expect(folderChildren).toHaveLength(1);
    expect(folderChildren[0]?.id).toBe("doc-nested");
  });

  test("uploads PDFs and moves them into the target folder", async () => {
    const moveCalls: Array<{ hash: string; parent: string }> = [];
    const connector = createRemarkableConnector(
      createMockApi({
        move: async (hash, parent) => {
          moveCalls.push({ hash, parent });
          return { hash: "hash-moved" };
        },
      }),
    );

    const uploaded = await connector.uploadDocument({
      name: "Brief.pdf",
      bytes: new Uint8Array([1, 2, 3]),
      mimeType: "application/pdf",
      parentId: "folder-1",
    });

    expect(uploaded).toMatchObject({
      id: "doc-2",
      hash: "hash-moved",
      name: "Brief.pdf",
      parentId: "folder-1",
      kind: "document",
      documentFormat: "pdf",
    });
    expect(moveCalls).toEqual([{ hash: "hash-doc-2", parent: "folder-1" }]);
  });

  test("downloads PDF bytes for pdf documents", async () => {
    const connector = createRemarkableConnector(createMockApi());
    const bytes = await connector.downloadDocument({
      id: "doc-1",
      hash: "hash-doc-1",
      name: "Contract.pdf",
      parentId: FOLIO_CONNECTOR_ROOT_ID,
      kind: "document",
      modifiedAt: "2026-07-01T12:00:00.000Z",
      starred: false,
      documentFormat: "pdf",
    });

    expect(bytes).toEqual(new Uint8Array([0x25, 0x50, 0x44, 0x46]));
  });
});
