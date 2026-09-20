import { describe, expect, test } from "bun:test";
import { Fragment } from "prosemirror-model";
import { initProseMirrorDoc, prosemirrorToYXmlFragment } from "y-prosemirror";
import * as Y from "yjs";

import { createEmptyDocument } from "../utils/createDocument";
import { mergeImageAttrs } from "./attrs";
import { toProseDoc } from "./conversion/toProseDoc";
import { schema } from "./schema";
import type { ImageAttrs } from "./schema/nodes";
import { SECTION_BREAK_TYPES } from "./sectionCarrier";
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

/** The last version that carried a section break on `sectionBreakType`. */
const LEGACY_ATTR_SCHEMA_VERSION = 3;
const LEGACY_PARAGRAPH_TEXT = "Ends the section";

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
  test("carries a snapshot from before it forward without rewriting anything", () => {
    const ydoc = new Y.Doc();
    const document = toProseDoc(createEmptyDocument({ initialText: "Pre-wrapper" }));
    const fragment = ydoc.getXmlFragment(PROSEMIRROR_FRAGMENT_NAME);
    prosemirrorToYXmlFragment(document, fragment);
    ydoc.getMap(METADATA_MAP_NAME).set(ATTR_SCHEMA_VERSION_KEY, 2);

    expect(applyAttrSchemaMigrations(ydoc, fragment, 2)).toBe(0);
    expect(initProseMirrorDoc(fragment, schema).doc.eq(document)).toBe(true);
    ydoc.destroy();
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

/**
 * The section break's one carrier, against snapshots written under version 3.
 *
 * A v3 paragraph could state `sectionBreakType` and nothing else, and the save
 * leg minted a `SectionProperties` from it. The attr is gone from the schema,
 * and ProseMirror drops an attr the schema does not declare without raising
 * anything, so an unmigrated v3 paragraph loads as one that ends no section:
 * the marker is what makes the loss visible, and the step is what repairs it.
 */
describe("the section break's one carrier against stored snapshots", () => {
  /** A v3 paragraph, built through Yjs because the schema no longer admits the attr. */
  const legacySnapshot = (attributes: Record<string, unknown>) => {
    const ydoc = new Y.Doc();
    const fragment = ydoc.getXmlFragment(PROSEMIRROR_FRAGMENT_NAME);
    const paragraph = new Y.XmlElement("paragraph");
    paragraph.insert(0, [new Y.XmlText(LEGACY_PARAGRAPH_TEXT)]);
    for (const [name, value] of Object.entries(attributes)) {
      // SAFETY: a Yjs attribute holds JSON; only the typings say `string`.
      paragraph.setAttribute(name, value as string);
    }
    fragment.insert(0, [paragraph]);
    ydoc.getMap(METADATA_MAP_NAME).set(ATTR_SCHEMA_VERSION_KEY, LEGACY_ATTR_SCHEMA_VERSION);
    return { fragment, ydoc };
  };

  const loadedParagraph = (ydoc: Y.Doc) =>
    initProseMirrorDoc(ydoc.getXmlFragment(PROSEMIRROR_FRAGMENT_NAME), schema).doc.child(0);

  const referenceParagraph = (properties: Record<string, unknown>) =>
    schema.node("paragraph", { _sectionProperties: properties }, [
      schema.text(LEGACY_PARAGRAPH_TEXT),
    ]);

  test.each(SECTION_BREAK_TYPES)("mints the record a v3 %s break meant", (breakType) => {
    const { fragment, ydoc } = legacySnapshot({ sectionBreakType: breakType });

    expect(applyAttrSchemaMigrations(ydoc, fragment, LEGACY_ATTR_SCHEMA_VERSION)).toBe(1);

    expect(loadedParagraph(ydoc).eq(referenceParagraph({ sectionStart: breakType }))).toBe(true);
    expect(readYjsAttrSchemaVersion(ydoc).isOk()).toBe(true);
    ydoc.destroy();
  });

  test("an unmigrated v3 paragraph loses the section the step exists to keep", () => {
    const { ydoc } = legacySnapshot({ sectionBreakType: "nextPage" });

    // Read without the step: the attr the schema no longer declares is dropped,
    // and with it the only statement that a section ended here.
    expect(loadedParagraph(ydoc).attrs["_sectionProperties"]).toBeNull();
    expect(loadedParagraph(ydoc).eq(referenceParagraph({ sectionStart: "nextPage" }))).toBe(false);
    ydoc.destroy();
  });

  test("a v3 paragraph that also held a record keeps the record, not the type", () => {
    const parsed = { sectionStart: "continuous", pageWidth: 11_906 };
    const { fragment, ydoc } = legacySnapshot({
      sectionBreakType: "nextPage",
      _sectionProperties: parsed,
    });

    expect(applyAttrSchemaMigrations(ydoc, fragment, LEGACY_ATTR_SCHEMA_VERSION)).toBe(1);

    // The record was always the authority the save leg preferred; a type beside
    // it that disagreed never reached the package.
    expect(loadedParagraph(ydoc).eq(referenceParagraph(parsed))).toBe(true);
    ydoc.destroy();
  });

  test("a v3 break type outside ST_SectionMark is dropped, not minted", () => {
    const { fragment, ydoc } = legacySnapshot({ sectionBreakType: "nextFrame" });

    expect(applyAttrSchemaMigrations(ydoc, fragment, LEGACY_ATTR_SCHEMA_VERSION)).toBe(1);

    // No producer could state it and no reader could resolve it, so there is no
    // section to keep.
    expect(loadedParagraph(ydoc).attrs["_sectionProperties"]).toBeNull();
    ydoc.destroy();
  });

  test("running the step again rewrites nothing", () => {
    const { fragment, ydoc } = legacySnapshot({ sectionBreakType: "oddPage" });
    applyAttrSchemaMigrations(ydoc, fragment, LEGACY_ATTR_SCHEMA_VERSION);
    const migrated = loadedParagraph(ydoc);

    expect(applyAttrSchemaMigrations(ydoc, fragment, LEGACY_ATTR_SCHEMA_VERSION)).toBe(0);

    expect(loadedParagraph(ydoc).eq(migrated)).toBe(true);
    ydoc.destroy();
  });

  test("the immediately preceding version carries exactly the new step", () => {
    expect(attrSchemaMigrationSteps(8)).toHaveLength(1);
    expect(attrSchemaMigrationSteps(0)).toHaveLength(FOLIO_YJS_ATTR_SCHEMA_VERSION);
  });
});
