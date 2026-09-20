import { describe, expect, test } from "bun:test";
import { prosemirrorToYXmlFragment } from "y-prosemirror";
import * as Y from "yjs";

import { toProseDoc } from "../../prosemirror/conversion/toProseDoc";
import {
  FOLIO_YJS_ATTR_SCHEMA_VERSION,
  applyAttrSchemaMigrations,
  readYjsAttrSchemaVersion,
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

/**
 * A version-3 snapshot holding the mark version 4 renames, spelled the way a
 * version-3 build persisted it: a Y.Text delta attribute keyed by the mark's
 * own name.
 *
 * Built here rather than committed as bytes, following the version-1 fixture
 * above: a reviewer can read what the old shape was, and the fixture does not
 * depend on a build that no longer exists.
 */
const versionThreeRunOwnerSnapshot = (): Uint8Array => {
  const ydoc = new Y.Doc();
  const text = new Y.XmlText();
  text.insert(0, "before");
  text.insert(6, "owned", { pageBreakRunOwner: { id: 12 } });
  // Stated rather than left to inherit: Yjs gives an insertion the attributes
  // of the character before it, which is the same edge `inclusive: false`
  // settles in the editor.
  text.insert(11, "after", { pageBreakRunOwner: null });
  const paragraph = new Y.XmlElement("paragraph");
  paragraph.insert(0, [text]);
  ydoc.getXmlFragment(FOLIO_YJS_PROSEMIRROR_FRAGMENT_NAME).insert(0, [paragraph]);
  ydoc.getMap(METADATA_MAP_NAME).set(ATTR_SCHEMA_VERSION_KEY, 3);
  const update = Y.encodeStateAsUpdate(ydoc);
  ydoc.destroy();
  return update;
};

type TextDeltaOp = { insert: string; attributes?: Record<string, unknown> };

const paragraphTextDelta = (update: Uint8Array): TextDeltaOp[] => {
  const ydoc = new Y.Doc();
  Y.applyUpdate(ydoc, update);
  const paragraph = ydoc.getXmlFragment(FOLIO_YJS_PROSEMIRROR_FRAGMENT_NAME).get(0);
  if (!(paragraph instanceof Y.XmlElement)) {
    throw new Error("Expected a paragraph element");
  }
  const text = paragraph.get(0);
  if (!(text instanceof Y.XmlText)) {
    throw new Error("Expected a shared text node");
  }
  const delta = text.toDelta() as TextDeltaOp[];
  ydoc.destroy();
  return delta;
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

/**
 * The rename is the first step that rewrites a *mark* attr, and its failure
 * mode is not a lost mark: `createTextNodesFromYText` wraps the rebuild of one
 * in a `try` whose `catch` deletes the Y.Text item, so an unknown mark name
 * costs the text under it.
 */
describe("migrateFolioYjsSnapshot renames a version-3 run owner mark", () => {
  test("the mark is renamed and the text it covered is untouched", () => {
    const migrated = migrateFolioYjsSnapshot(versionThreeRunOwnerSnapshot());
    if (migrated.isErr()) {
      throw migrated.error;
    }

    expect(migrated.value.fromVersion).toBe(3);
    expect(migrated.value.toVersion).toBe(FOLIO_YJS_ATTR_SCHEMA_VERSION);
    const delta = paragraphTextDelta(migrated.value.update);

    expect(delta.map(({ insert }) => insert).join("")).toBe("beforeownedafter");
    const owned = delta.find(({ insert }) => insert === "owned");
    expect(owned?.attributes?.["pageBreakRunOwner"]).toBeUndefined();
    expect(owned?.attributes?.["runIdentity"]).toEqual({ id: 12 });
    for (const op of delta.filter(({ insert }) => insert !== "owned")) {
      expect(op.attributes?.["runIdentity"]).toBeUndefined();
    }
  });

  test("a snapshot with no such mark is left byte for byte alone", () => {
    const ydoc = new Y.Doc();
    Y.applyUpdate(ydoc, snapshot());
    const before = Y.encodeStateAsUpdate(ydoc);
    ydoc.getMap(METADATA_MAP_NAME).set(ATTR_SCHEMA_VERSION_KEY, 3);
    const stale = Y.encodeStateAsUpdate(ydoc);
    ydoc.destroy();

    const migrated = migrateFolioYjsSnapshot(stale);
    if (migrated.isErr()) {
      throw migrated.error;
    }
    expect(migrated.value.paragraphsRewritten).toBe(0);

    const after = new Y.Doc();
    Y.applyUpdate(after, migrated.value.update);
    const original = new Y.Doc();
    Y.applyUpdate(original, before);
    expect(after.getXmlFragment(FOLIO_YJS_PROSEMIRROR_FRAGMENT_NAME).toJSON()).toBe(
      original.getXmlFragment(FOLIO_YJS_PROSEMIRROR_FRAGMENT_NAME).toJSON(),
    );
    after.destroy();
    original.destroy();
  });

  test("a build that predates the rename refuses the migrated snapshot", () => {
    // The reverse direction, and the reason the marker moves at all: without
    // it an older build would read a `runIdentity` it does not know and take
    // the same deleting path this step exists to avoid.
    const error = expectError(snapshot(FOLIO_YJS_ATTR_SCHEMA_VERSION + 1));

    expect(error.code).toBe("unsupported_version");
  });
});

/**
 * A load path runs the steps too. `migrateFolioYjsSnapshot` is the offline
 * sweep; an editor and a materialization read a stored snapshot directly, and
 * a step that rewrites values has to run before a node is built from them.
 */
describe("a load path carries an older fragment forward", () => {
  test("the steps run and the marker is stamped before anything reads the fragment", () => {
    const ydoc = new Y.Doc();
    Y.applyUpdate(ydoc, versionOneFieldSnapshot());
    const fragment = ydoc.getXmlFragment(FOLIO_YJS_PROSEMIRROR_FRAGMENT_NAME);

    const before = readYjsAttrSchemaVersion(ydoc);
    expect(before.isOk() ? before.value : null).toBe(1);
    expect(applyAttrSchemaMigrations(ydoc, fragment, 1)).toBe(1);

    const after = readYjsAttrSchemaVersion(ydoc);
    expect(after.isOk() ? after.value : null).toBe(FOLIO_YJS_ATTR_SCHEMA_VERSION);
    const attributes = fieldAttributes(Y.encodeStateAsUpdate(ydoc));
    expect(attributes["fldLock"]).toBeUndefined();
    expect(attributes["dirty"]).toBeUndefined();
    ydoc.destroy();
  });

  test("a fragment already at the current version is left alone", () => {
    const ydoc = new Y.Doc();
    Y.applyUpdate(ydoc, snapshot());
    const before = Y.encodeStateAsUpdate(ydoc);

    expect(
      applyAttrSchemaMigrations(
        ydoc,
        ydoc.getXmlFragment(FOLIO_YJS_PROSEMIRROR_FRAGMENT_NAME),
        FOLIO_YJS_ATTR_SCHEMA_VERSION,
      ),
    ).toBe(0);

    expect(Y.encodeStateAsUpdate(ydoc)).toEqual(before);
    ydoc.destroy();
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
