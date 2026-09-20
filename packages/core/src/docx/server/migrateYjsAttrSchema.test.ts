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
 * A version-3 snapshot holding three paragraphs: a heading level, the reserved
 * body-text nine, and a value outside the range a real package still carries.
 */
const versionThreeOutlineSnapshot = (): Uint8Array => {
  const ydoc = new Y.Doc();
  const paragraphs = [0, 9, 12].map((stated) => {
    const paragraph = new Y.XmlElement("paragraph");
    // @ts-expect-error — a Yjs attribute holds JSON, and the version-3 shape
    // (the `w:outlineLvl w:val` number) is the point of the test.
    paragraph.setAttribute("outlineLevel", stated);
    return paragraph;
  });
  ydoc.getXmlFragment(FOLIO_YJS_PROSEMIRROR_FRAGMENT_NAME).insert(0, paragraphs);
  ydoc.getMap(METADATA_MAP_NAME).set(ATTR_SCHEMA_VERSION_KEY, 3);
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

describe("migrateFolioYjsSnapshot carries a version-3 outline level forward", () => {
  test("maps the stated number onto the union and drops what the format never defined", () => {
    const migrated = migrateFolioYjsSnapshot(versionThreeOutlineSnapshot());
    if (migrated.isErr()) {
      throw migrated.error;
    }

    expect(migrated.value.fromVersion).toBe(3);
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
  test("an unmigrated version-3 paragraph is refused rather than misread", () => {
    const ydoc = new Y.Doc();
    Y.applyUpdate(ydoc, versionThreeOutlineSnapshot());
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

/**
 * A version-4 snapshot holding the two `<w:numPr>` slots every way a version-4
 * build could have stored them, including the divergence the version-5 step
 * exists to end: `_originalFormatting` already carried the model's union while
 * the attr beside it carried the slots.
 */
const versionFourNumberingSnapshot = (): Uint8Array => {
  const ydoc = new Y.Doc();
  const stored: readonly Record<string, unknown>[] = [
    { numPr: { numId: 3, ilvl: 2 } },
    { numPr: { numId: 3 } },
    { numPr: { ilvl: 1 } },
    { numPr: { numId: 0, ilvl: 4 } },
    { numPr: {} },
    { numPr: { numId: 7 }, numPrFromStyle: { numId: 7 } },
    {
      _propertyChanges: [
        {
          type: "paragraphPropertyChange",
          info: { id: 1, author: "Reviewer", date: "2026-01-01" },
          previousFormatting: { numPr: null },
        },
        {
          type: "paragraphPropertyChange",
          info: { id: 2, author: "Reviewer", date: "2026-01-01" },
          previousFormatting: { numPr: { numId: 9, ilvl: 0 } },
        },
      ],
      _originalFormatting: { numPr: { kind: "reference", numId: 9 } },
      numPr: { numId: 9, ilvl: 0 },
    },
    { styleId: "Normal" },
  ];
  const paragraphs = stored.map((attributes) => {
    const paragraph = new Y.XmlElement("paragraph");
    for (const [key, value] of Object.entries(attributes)) {
      // @ts-expect-error — a Yjs attribute holds JSON, and the version-4 shape
      // is the point of the test; the typings narrow to string.
      paragraph.setAttribute(key, value);
    }
    return paragraph;
  });
  ydoc.getXmlFragment(FOLIO_YJS_PROSEMIRROR_FRAGMENT_NAME).insert(0, paragraphs);
  ydoc.getMap(METADATA_MAP_NAME).set(ATTR_SCHEMA_VERSION_KEY, 4);
  const update = Y.encodeStateAsUpdate(ydoc);
  ydoc.destroy();
  return update;
};

const paragraphAttributes = (update: Uint8Array): Record<string, unknown>[] => {
  const ydoc = new Y.Doc();
  Y.applyUpdate(ydoc, update);
  const attributes = ydoc
    .getXmlFragment(FOLIO_YJS_PROSEMIRROR_FRAGMENT_NAME)
    .toArray()
    .map((node) => {
      if (!(node instanceof Y.XmlElement)) {
        throw new Error("Expected a paragraph element");
      }
      const record: Record<string, unknown> = node.getAttributes();
      return record;
    });
  ydoc.destroy();
  return attributes;
};

describe("migrateFolioYjsSnapshot carries version-4 numbering forward", () => {
  test("maps the two slots onto the union, everywhere a version-4 build stored them", () => {
    const migrated = migrateFolioYjsSnapshot(versionFourNumberingSnapshot());
    if (migrated.isErr()) {
      throw migrated.error;
    }

    expect(migrated.value.fromVersion).toBe(4);
    expect(migrated.value.toVersion).toBe(FOLIO_YJS_ATTR_SCHEMA_VERSION);
    expect(migrated.value.paragraphsRewritten).toBe(7);
    expect(paragraphAttributes(migrated.value.update)).toEqual([
      { numPr: { kind: "reference", numId: 3, ilvl: 2 } },
      { numPr: { kind: "reference", numId: 3 } },
      { numPr: { kind: "levelOnly", ilvl: 1 } },
      // A cancellation names no id, so the level it sat beside goes with it.
      { numPr: { kind: "none" } },
      {},
      {
        numPr: { kind: "reference", numId: 7 },
        numPrFromStyle: { kind: "reference", numId: 7 },
      },
      {
        _propertyChanges: [
          // `null` is the tombstone for "carried no numbering", not a slot pair.
          {
            type: "paragraphPropertyChange",
            info: { id: 1, author: "Reviewer", date: "2026-01-01" },
            previousFormatting: { numPr: null },
          },
          {
            type: "paragraphPropertyChange",
            info: { id: 2, author: "Reviewer", date: "2026-01-01" },
            previousFormatting: { numPr: { kind: "reference", numId: 9, ilvl: 0 } },
          },
        ],
        // Already a union, and left exactly as it was found.
        _originalFormatting: { numPr: { kind: "reference", numId: 9 } },
        numPr: { kind: "reference", numId: 9, ilvl: 0 },
      },
      { styleId: "Normal" },
    ]);
  });

  /**
   * The gate has to fire before anything reads the fragment. Without it the
   * slot pair reaches the node verbatim and every `switch` over the union
   * reads `kind === undefined` as the arm it is not — a numbered paragraph
   * that quietly stops being numbered, and the loss written back on the next
   * debounce.
   */
  test("an unmigrated version-4 paragraph is refused rather than misread", () => {
    const ydoc = new Y.Doc();
    Y.applyUpdate(ydoc, versionFourNumberingSnapshot());
    const document = initProseMirrorDoc(
      ydoc.getXmlFragment(FOLIO_YJS_PROSEMIRROR_FRAGMENT_NAME),
      schema,
    ).doc;
    ydoc.destroy();

    const result = readParagraphAttrs(document.child(0));
    expect(result.ok).toBe(false);
    expect(result.ok ? [] : result.issues.map((issue) => issue.path)).toEqual([
      "paragraph.attrs.numPr",
    ]);
  });

  test("a migrated version-4 paragraph reads", () => {
    const migrated = migrateFolioYjsSnapshot(versionFourNumberingSnapshot());
    if (migrated.isErr()) {
      throw migrated.error;
    }
    const ydoc = new Y.Doc();
    Y.applyUpdate(ydoc, migrated.value.update);
    const document = initProseMirrorDoc(
      ydoc.getXmlFragment(FOLIO_YJS_PROSEMIRROR_FRAGMENT_NAME),
      schema,
    ).doc;
    ydoc.destroy();

    for (let index = 0; index < document.childCount; index += 1) {
      expect(readParagraphAttrs(document.child(index)).ok).toBe(true);
    }
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
