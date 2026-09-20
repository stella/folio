import { describe, expect, test } from "bun:test";
import { prosemirrorToYXmlFragment } from "y-prosemirror";
import * as Y from "yjs";

import { toProseDoc } from "../../prosemirror/conversion/toProseDoc";
import {
  FOLIO_YJS_ATTR_SCHEMA_VERSION,
  writeYjsDocumentMetadata,
} from "../../prosemirror/yjsDocumentMetadata";
import { createEmptyDocument } from "../../utils/createDocument";
import {
  FOLIO_YJS_PROSEMIRROR_FRAGMENT_NAME,
  FOLIO_YJS_UPDATE_MAX_BYTES,
} from "./materializeYjsDocx";
import { FolioYjsSnapshotMigrationError, migrateFolioYjsSnapshot } from "./migrateYjsAttrSchema";

const METADATA_MAP_NAME = "folio:document-metadata";
const ATTR_SCHEMA_VERSION_KEY = "attrSchemaVersion";

const snapshot = (marker?: unknown): Uint8Array => {
  const ydoc = new Y.Doc();
  const document = toProseDoc(createEmptyDocument({ initialText: "Room" }));
  prosemirrorToYXmlFragment(document, ydoc.getXmlFragment(FOLIO_YJS_PROSEMIRROR_FRAGMENT_NAME));
  writeYjsDocumentMetadata(ydoc, document);
  if (marker !== undefined) {
    ydoc.getMap(METADATA_MAP_NAME).set(ATTR_SCHEMA_VERSION_KEY, marker);
  }
  const update = Y.encodeStateAsUpdate(ydoc);
  ydoc.destroy();
  return update;
};

/**
 * A version-1 snapshot holding one field, with the flags a version-1 build
 * stored: `false` for a field that authored nothing, because the reader that
 * filled them tested `=== true`.
 */
const versionOneFieldSnapshot = (): Uint8Array => {
  const ydoc = new Y.Doc();
  const field = new Y.XmlElement("field");
  field.setAttribute("fieldType", "PAGE");
  field.setAttribute("instruction", " PAGE ");
  field.setAttribute("displayText", "");
  field.setAttribute("fieldKind", "complex");
  // @ts-expect-error — a Yjs attribute holds JSON, and the stored shape is the
  // point of the test; the typings narrow to string.
  field.setAttribute("fldLock", false);
  // @ts-expect-error — as above.
  field.setAttribute("dirty", false);
  const paragraph = new Y.XmlElement("paragraph");
  paragraph.insert(0, [field]);
  ydoc.getXmlFragment(FOLIO_YJS_PROSEMIRROR_FRAGMENT_NAME).insert(0, [paragraph]);
  ydoc.getMap(METADATA_MAP_NAME).set(ATTR_SCHEMA_VERSION_KEY, 1);
  const update = Y.encodeStateAsUpdate(ydoc);
  ydoc.destroy();
  return update;
};

const fieldAttributes = (update: Uint8Array): Record<string, unknown> => {
  const ydoc = new Y.Doc();
  Y.applyUpdate(ydoc, update);
  const paragraph = ydoc.getXmlFragment(FOLIO_YJS_PROSEMIRROR_FRAGMENT_NAME).get(0);
  if (!(paragraph instanceof Y.XmlElement)) {
    throw new Error("Expected a paragraph element");
  }
  const field = paragraph.get(0);
  if (!(field instanceof Y.XmlElement)) {
    throw new Error("Expected a field element");
  }
  const attributes = field.getAttributes();
  ydoc.destroy();
  return attributes;
};

const expectError = (update: Uint8Array): FolioYjsSnapshotMigrationError => {
  const migrated = migrateFolioYjsSnapshot(update);
  if (migrated.isOk()) {
    throw new Error("Expected the migration to fail");
  }
  return migrated.error;
};

describe("migrateFolioYjsSnapshot carries a version-1 field forward", () => {
  test("drops the field flags version 1 could not have stated", () => {
    const migrated = migrateFolioYjsSnapshot(versionOneFieldSnapshot());
    if (migrated.isErr()) {
      throw migrated.error;
    }

    expect(migrated.value.fromVersion).toBe(1);
    expect(migrated.value.toVersion).toBe(FOLIO_YJS_ATTR_SCHEMA_VERSION);
    expect(migrated.value.paragraphsRewritten).toBe(1);
    const attributes = fieldAttributes(migrated.value.update);
    expect(attributes["fldLock"]).toBeUndefined();
    expect(attributes["dirty"]).toBeUndefined();
    expect(attributes["fieldType"]).toBe("PAGE");
  });
});

describe("migrateFolioYjsSnapshot failures", () => {
  test("refuses a snapshot written by newer code", () => {
    const error = expectError(snapshot(FOLIO_YJS_ATTR_SCHEMA_VERSION + 1));

    expect(error).toBeInstanceOf(FolioYjsSnapshotMigrationError);
    expect(error.code).toBe("unsupported_version");
  });

  test("refuses an empty update", () => {
    expect(expectError(new Uint8Array()).code).toBe("invalid_update");
  });

  test("refuses an update that is not Yjs state", () => {
    expect(expectError(new Uint8Array([255, 255, 255, 255])).code).toBe("invalid_update");
  });

  test("refuses an update above the snapshot limit", () => {
    expect(expectError(new Uint8Array(FOLIO_YJS_UPDATE_MAX_BYTES + 1)).code).toBe(
      "update_too_large",
    );
  });
});
