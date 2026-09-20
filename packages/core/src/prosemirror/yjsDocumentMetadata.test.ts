import { describe, expect, test } from "bun:test";
import { Fragment } from "prosemirror-model";
import { initProseMirrorDoc, prosemirrorToYXmlFragment } from "y-prosemirror";
import * as Y from "yjs";

import { createEmptyDocument } from "../utils/createDocument";
import { mergeImageAttrs } from "./attrs";
import { toProseDoc } from "./conversion/toProseDoc";
import { schema } from "./schema";
import type { ImageAttrs } from "./schema/nodes";
import {
  applyAttrSchemaMigrations,
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

/**
 * The `inlineWrapper` mark adds one attr whose default is `null`, and no
 * snapshot written before it states the attr. `null` is the absence of a
 * wrapper, which is what those snapshots meant, so there is nothing for a
 * migration step to rewrite and the marker must not move: bumping it would
 * make every existing snapshot look stale and make an older build refuse one
 * this build wrote for no reason.
 */
describe("the inline wrapper mark against stored snapshots", () => {
  test("does not move the attr-schema version", () => {
    expect(FOLIO_YJS_ATTR_SCHEMA_VERSION).toBe(4);
    expect(attrSchemaMigrationSteps(FOLIO_YJS_ATTR_SCHEMA_VERSION)).toHaveLength(0);
  });

  test("a snapshot that states no stack loads as a document with no wrapper", () => {
    const { document, ydoc } = seededDocument();
    const loaded = initProseMirrorDoc(ydoc.getXmlFragment(PROSEMIRROR_FRAGMENT_NAME), schema).doc;

    expect(loaded.eq(document)).toBe(true);
    loaded.descendants((node) => {
      expect(node.marks.map((mark) => mark.type.name)).not.toContain("inlineWrapper");
    });
    ydoc.destroy();
  });
});

/**
 * A version 3 snapshot holds a VML shape's render as an ordinary editable
 * picture that states a relationship it does not have. Carrying it forward has
 * to answer both: the attr loses the spelling no relationship answers to, and
 * the node is classified so the next edit keeps the capture instead of
 * dropping it and saving the render in the shape's place.
 */
describe("a stored preview carried forward from version 3", () => {
  const V3_CAPTURE =
    '<w:pict><v:rect style="width:10pt;height:10pt" fillcolor="#abcdef"/></w:pict>';

  const storedImageElement = (fragment: Y.XmlFragment): Y.XmlElement => {
    const find = (node: Y.XmlElement | Y.XmlFragment): Y.XmlElement | undefined => {
      if ("nodeName" in node && node.nodeName === "image") {
        return node;
      }
      for (const child of node.toArray()) {
        if (typeof child === "string" || !("toArray" in child)) {
          continue;
        }
        const found = find(child);
        if (found) {
          return found;
        }
      }
      return undefined;
    };
    const image = find(fragment);
    if (!image) {
      throw new Error("the snapshot holds no image node");
    }
    return image;
  };

  /** The attrs a version 3 build wrote for a VML shape it rendered. */
  const versionThreeSnapshot = (): Y.Doc => {
    const ydoc = new Y.Doc();
    const document = toProseDoc(createEmptyDocument({ initialText: "Shape" }));
    const paragraph = document.child(0);
    const image = schema.nodes["image"]!.create({
      src: "data:image/svg+xml;charset=utf-8,%3Csvg%3E",
      // oxlint-disable-next-line folio-relationship-ids/no-empty-relationship-id -- the spelling version 3 stored is what this step has to read
      rId: "",
      width: 13,
      height: 13,
      _docxRawXml: V3_CAPTURE,
    });
    const withImage = document.copy(
      document.content.replaceChild(0, paragraph.copy(Fragment.from(image))),
    );
    prosemirrorToYXmlFragment(withImage, ydoc.getXmlFragment(PROSEMIRROR_FRAGMENT_NAME));
    return ydoc;
  };

  test("loses the empty relationship and keeps the markup an edit used to drop", () => {
    const ydoc = versionThreeSnapshot();
    const fragment = ydoc.getXmlFragment(PROSEMIRROR_FRAGMENT_NAME);

    expect(applyAttrSchemaMigrations(ydoc, fragment, 3)).toBe(1);

    const migrated: Record<string, unknown> = storedImageElement(fragment).getAttributes();
    expect(migrated["rId"]).toBeUndefined();
    expect(migrated["_docxRawXmlMode"]).toBe("previewOnly");
    expect(migrated["_docxRawXml"]).toBe(V3_CAPTURE);

    // What the classification buys: the edit that used to drop the capture
    // leaves it, so the save writes the shape rather than the render.
    const loaded = initProseMirrorDoc(fragment, schema).doc;
    let edited: ImageAttrs | undefined;
    loaded.descendants((node) => {
      if (node.type.name === "image") {
        edited = mergeImageAttrs(node, { width: 200 });
      }
      return true;
    });
    expect(edited?._docxRawXml).toBe(V3_CAPTURE);
    expect(edited?.rId).toBeUndefined();
    ydoc.destroy();
  });

  test("leaves a drawing that names a relationship alone", () => {
    const ydoc = versionThreeSnapshot();
    const fragment = ydoc.getXmlFragment(PROSEMIRROR_FRAGMENT_NAME);
    storedImageElement(fragment).setAttribute("rId", "rId7");

    expect(applyAttrSchemaMigrations(ydoc, fragment, 3)).toBe(0);

    const migrated: Record<string, unknown> = storedImageElement(fragment).getAttributes();
    expect(migrated["rId"]).toBe("rId7");
    expect(migrated["_docxRawXmlMode"]).toBeUndefined();
    ydoc.destroy();
  });
});
