import { describe, expect, test } from "bun:test";
import { initProseMirrorDoc, prosemirrorToYXmlFragment } from "y-prosemirror";
import * as Y from "yjs";

import { createEmptyDocument } from "../utils/createDocument";
import { toProseDoc } from "./conversion/toProseDoc";
import { schema } from "./schema";
import {
  FOLIO_YJS_ATTR_SCHEMA_VERSION,
  FolioYjsAttrSchemaVersionError,
  attrSchemaMigrationSteps,
  readYjsAttrSchemaVersion,
  writeYjsDocumentMetadata,
} from "./yjsDocumentMetadata";

const METADATA_MAP_NAME = "folio:document-metadata";
const ATTR_SCHEMA_VERSION_KEY = "attrSchemaVersion";
const PROSEMIRROR_FRAGMENT_NAME = "prosemirror";

const seededDocument = () => {
  const ydoc = new Y.Doc();
  const document = toProseDoc(createEmptyDocument({ initialText: "Seeded" }));
  prosemirrorToYXmlFragment(document, ydoc.getXmlFragment(PROSEMIRROR_FRAGMENT_NAME));
  writeYjsDocumentMetadata(ydoc, document);
  return { document, ydoc };
};

describe("attr-schema version marker", () => {
  test("seeding writes the marker this build reads", () => {
    const { ydoc } = seededDocument();

    expect(ydoc.getMap(METADATA_MAP_NAME).get(ATTR_SCHEMA_VERSION_KEY)).toBe(
      FOLIO_YJS_ATTR_SCHEMA_VERSION,
    );
    const read = readYjsAttrSchemaVersion(ydoc);
    expect(read.isOk() ? read.value : null).toBe(FOLIO_YJS_ATTR_SCHEMA_VERSION);
    ydoc.destroy();
  });

  test("an unversioned snapshot reads as version 0 and still loads", () => {
    const ydoc = new Y.Doc();
    const document = toProseDoc(createEmptyDocument({ initialText: "Pre-marker" }));
    prosemirrorToYXmlFragment(document, ydoc.getXmlFragment(PROSEMIRROR_FRAGMENT_NAME));

    const read = readYjsAttrSchemaVersion(ydoc);
    expect(read.isOk() ? read.value : null).toBe(0);
    expect(
      initProseMirrorDoc(ydoc.getXmlFragment(PROSEMIRROR_FRAGMENT_NAME), schema).doc.eq(document),
    ).toBe(true);
    ydoc.destroy();
  });

  test("a marker from newer code is refused instead of read", () => {
    const { ydoc } = seededDocument();
    ydoc.getMap(METADATA_MAP_NAME).set(ATTR_SCHEMA_VERSION_KEY, FOLIO_YJS_ATTR_SCHEMA_VERSION + 1);

    const read = readYjsAttrSchemaVersion(ydoc);
    expect(read.isErr()).toBe(true);
    if (read.isErr()) {
      expect(read.error).toBeInstanceOf(FolioYjsAttrSchemaVersionError);
      expect(read.error.marker).toBe(FOLIO_YJS_ATTR_SCHEMA_VERSION + 1);
      expect(read.error.supportedVersion).toBe(FOLIO_YJS_ATTR_SCHEMA_VERSION);
    }
    ydoc.destroy();
  });

  test.each([
    ["a string spelling of the current version", String(FOLIO_YJS_ATTR_SCHEMA_VERSION)],
    ["a fractional version", 1.5],
    ["a negative version", -1],
  ])("refuses %s", (_name, marker) => {
    const { ydoc } = seededDocument();
    ydoc.getMap(METADATA_MAP_NAME).set(ATTR_SCHEMA_VERSION_KEY, marker);

    expect(readYjsAttrSchemaVersion(ydoc).isErr()).toBe(true);
    ydoc.destroy();
  });

  test("the metadata write survives a y-prosemirror round trip", () => {
    const { document, ydoc } = seededDocument();
    const update = Y.encodeStateAsUpdate(ydoc);
    ydoc.destroy();

    const reloaded = new Y.Doc();
    Y.applyUpdate(reloaded, update);
    const read = readYjsAttrSchemaVersion(reloaded);

    expect(read.isOk() ? read.value : null).toBe(FOLIO_YJS_ATTR_SCHEMA_VERSION);
    expect(
      initProseMirrorDoc(reloaded.getXmlFragment(PROSEMIRROR_FRAGMENT_NAME), schema).doc.eq(
        document,
      ),
    ).toBe(true);
    reloaded.destroy();
  });

  test("every version below the current one carries a migration step", () => {
    expect(attrSchemaMigrationSteps(FOLIO_YJS_ATTR_SCHEMA_VERSION)).toHaveLength(0);
    expect(attrSchemaMigrationSteps(0)).toHaveLength(FOLIO_YJS_ATTR_SCHEMA_VERSION);
  });
});
