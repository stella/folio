import { describe, expect, test } from "bun:test";
import { initProseMirrorDoc, prosemirrorToYXmlFragment } from "y-prosemirror";
import * as Y from "yjs";

import { readParagraphAttrs } from "../../prosemirror/attrs";
import { schema } from "../../prosemirror/schema";

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
 * A version-4 snapshot holding one row of two cells, with the widths a
 * version-4 build stored: the cell that stated a `w:tcW` and the one that
 * states none both carry `width`, because it is the width the table resolved.
 */
const versionFourTableSnapshot = (): Uint8Array => {
  const ydoc = new Y.Doc();
  const cells = [
    { width: 2400, original: { width: { value: 2400, type: "dxa" } } },
    { width: 2400, original: { vAlign: "center" } },
  ].map(({ width, original }) => {
    const cell = new Y.XmlElement("tableCell");
    // @ts-expect-error — a Yjs attribute holds JSON, and the stored shape is
    // the point of the test; the typings narrow to string.
    cell.setAttribute("width", width);
    cell.setAttribute("widthType", "dxa");
    // @ts-expect-error — as above.
    cell.setAttribute("_originalFormatting", original);
    cell.insert(0, [new Y.XmlElement("paragraph")]);
    return cell;
  });
  const row = new Y.XmlElement("tableRow");
  row.insert(0, cells);
  const table = new Y.XmlElement("table");
  table.insert(0, [row]);
  ydoc.getXmlFragment(FOLIO_YJS_PROSEMIRROR_FRAGMENT_NAME).insert(0, [table]);
  ydoc.getMap(METADATA_MAP_NAME).set(ATTR_SCHEMA_VERSION_KEY, 4);
  const update = Y.encodeStateAsUpdate(ydoc);
  ydoc.destroy();
  return update;
};

/**
 * A version-5 snapshot holding one row, with the `hidden` a version-5 build
 * stored: `false` for a row that authored no `w:hidden`, because the reader
 * that filled it took only an explicit on.
 */
const versionFiveRowSnapshot = (): Uint8Array => {
  const ydoc = new Y.Doc();
  const cell = new Y.XmlElement("tableCell");
  cell.insert(0, [new Y.XmlElement("paragraph")]);
  const row = new Y.XmlElement("tableRow");
  // @ts-expect-error — a Yjs attribute holds JSON, and the stored shape is the
  // point of the test; the typings narrow to string.
  row.setAttribute("hidden", false);
  // @ts-expect-error — as above.
  row.setAttribute("isHeader", false);
  row.insert(0, [cell]);
  const table = new Y.XmlElement("table");
  table.insert(0, [row]);
  ydoc.getXmlFragment(FOLIO_YJS_PROSEMIRROR_FRAGMENT_NAME).insert(0, [table]);
  ydoc.getMap(METADATA_MAP_NAME).set(ATTR_SCHEMA_VERSION_KEY, 5);
  const update = Y.encodeStateAsUpdate(ydoc);
  ydoc.destroy();
  return update;
};

const rowAttributes = (update: Uint8Array): Record<string, unknown> => {
  const ydoc = new Y.Doc();
  Y.applyUpdate(ydoc, update);
  const table = ydoc.getXmlFragment(FOLIO_YJS_PROSEMIRROR_FRAGMENT_NAME).get(0);
  if (!(table instanceof Y.XmlElement)) {
    throw new Error("Expected a table element");
  }
  const row = table.get(0);
  if (!(row instanceof Y.XmlElement)) {
    throw new Error("Expected a row element");
  }
  const attributes = row.getAttributes();
  ydoc.destroy();
  return attributes;
};

const cellAttributes = (update: Uint8Array): Record<string, unknown>[] => {
  const ydoc = new Y.Doc();
  Y.applyUpdate(ydoc, update);
  const table = ydoc.getXmlFragment(FOLIO_YJS_PROSEMIRROR_FRAGMENT_NAME).get(0);
  if (!(table instanceof Y.XmlElement)) {
    throw new Error("Expected a table element");
  }
  const row = table.get(0);
  if (!(row instanceof Y.XmlElement)) {
    throw new Error("Expected a row element");
  }
  const attributes = row.toArray().map((cell) => {
    if (!(cell instanceof Y.XmlElement)) {
      throw new Error("Expected a cell element");
    }
    return cell.getAttributes();
  });
  ydoc.destroy();
  return attributes;
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

describe("migrateFolioYjsSnapshot carries a version-4 table forward", () => {
  test("backfills a stated width only onto the cell that stated one", () => {
    const migrated = migrateFolioYjsSnapshot(versionFourTableSnapshot());
    if (migrated.isErr()) {
      throw migrated.error;
    }

    expect(migrated.value.fromVersion).toBe(4);
    expect(migrated.value.paragraphsRewritten).toBe(1);
    const [stated, resolved] = cellAttributes(migrated.value.update);
    expect(stated?.["_authoredWidth"]).toEqual({ value: 2400, type: "dxa" });
    // The width this cell holds is the one the table's grid resolved, so a
    // save must not hand it back as a `w:tcW` its author never wrote.
    expect(resolved?.["_authoredWidth"]).toBeUndefined();
    expect(resolved?.["width"]).toBe(2400);
  });
});

describe("migrateFolioYjsSnapshot carries a version-5 row forward", () => {
  test("drops the `hidden` version 5 could not have stated", () => {
    const migrated = migrateFolioYjsSnapshot(versionFiveRowSnapshot());
    if (migrated.isErr()) {
      throw migrated.error;
    }

    expect(migrated.value.fromVersion).toBe(5);
    expect(migrated.value.paragraphsRewritten).toBe(1);
    const attributes = rowAttributes(migrated.value.update);
    expect(attributes["hidden"]).toBeUndefined();
    // `isHeader` still answers a permanent yes/no question, so its `false`
    // means what it says and the step leaves it alone.
    expect(attributes["isHeader"]).toBe(false);
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

/**
 * A version-6 snapshot holding three paragraphs: a heading level, the reserved
 * body-text nine, and a value outside the range a real package still carries.
 */
const versionSixOutlineSnapshot = (): Uint8Array => {
  const ydoc = new Y.Doc();
  const paragraphs = [0, 9, 12].map((stated) => {
    const paragraph = new Y.XmlElement("paragraph");
    // @ts-expect-error — a Yjs attribute holds JSON, and the version-6 shape
    // (the `w:outlineLvl w:val` number) is the point of the test.
    paragraph.setAttribute("outlineLevel", stated);
    return paragraph;
  });
  ydoc.getXmlFragment(FOLIO_YJS_PROSEMIRROR_FRAGMENT_NAME).insert(0, paragraphs);
  ydoc.getMap(METADATA_MAP_NAME).set(ATTR_SCHEMA_VERSION_KEY, 6);
  const update = Y.encodeStateAsUpdate(ydoc);
  ydoc.destroy();
  return update;
};

const paragraphOutlineLevels = (update: Uint8Array): unknown[] => {
  const ydoc = new Y.Doc();
  Y.applyUpdate(ydoc, update);
  const levels = ydoc
    .getXmlFragment(FOLIO_YJS_PROSEMIRROR_FRAGMENT_NAME)
    .toArray()
    .map((node) => {
      if (!(node instanceof Y.XmlElement)) {
        throw new Error("Expected a paragraph element");
      }
      const attributes: Record<string, unknown> = node.getAttributes();
      return attributes["outlineLevel"];
    });
  ydoc.destroy();
  return levels;
};

describe("migrateFolioYjsSnapshot carries a version-6 outline level forward", () => {
  test("maps the stated number onto the union and drops what the format never defined", () => {
    const migrated = migrateFolioYjsSnapshot(versionSixOutlineSnapshot());
    if (migrated.isErr()) {
      throw migrated.error;
    }

    expect(migrated.value.fromVersion).toBe(6);
    expect(migrated.value.toVersion).toBe(FOLIO_YJS_ATTR_SCHEMA_VERSION);
    expect(migrated.value.paragraphsRewritten).toBe(3);
    expect(paragraphOutlineLevels(migrated.value.update)).toEqual([
      { kind: "heading", level: 0 },
      { kind: "bodyText" },
      undefined,
    ]);
  });

  /**
   * The gate has to fire before anything reads the fragment. Without it the
   * stored number reaches the node verbatim, the strict validator refuses it,
   * and a room that opened yesterday stops opening — which is the recoverable
   * half of the failure. The unrecoverable half is what the old code did with
   * it: read `9` as a tenth heading level.
   */
  test("an unmigrated version-6 paragraph is refused rather than misread", () => {
    const ydoc = new Y.Doc();
    Y.applyUpdate(ydoc, versionSixOutlineSnapshot());
    const document = initProseMirrorDoc(
      ydoc.getXmlFragment(FOLIO_YJS_PROSEMIRROR_FRAGMENT_NAME),
      schema,
    ).doc;
    ydoc.destroy();

    const result = readParagraphAttrs(document.child(0));
    expect(result.ok).toBe(false);
    expect(result.ok ? [] : result.issues.map((issue) => issue.path)).toEqual([
      "paragraph.attrs.outlineLevel",
    ]);
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
