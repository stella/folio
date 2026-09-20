import { describe, expect, test } from "bun:test";
import { initProseMirrorDoc, prosemirrorToYXmlFragment } from "y-prosemirror";
import * as Y from "yjs";

import { createEmptyDocument } from "../utils/createDocument";
import { toProseDoc } from "./conversion/toProseDoc";
import { schema } from "./schema";
import { SECTION_BREAK_TYPES } from "./sectionCarrier";
import {
  FOLIO_YJS_ATTR_SCHEMA_VERSION,
  FolioYjsAttrSchemaVersionError,
  applyAttrSchemaMigrations,
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
  test("does not move the attr-schema version", () => {
    // The pin is the point: a bump must be a decision someone made about the
    // attrs it carries, not one a wrapper mark collected on its way past.
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

  test("a v3 break type the editor never authored is dropped, not minted", () => {
    const { fragment, ydoc } = legacySnapshot({ sectionBreakType: "nextColumn" });

    expect(applyAttrSchemaMigrations(ydoc, fragment, LEGACY_ATTR_SCHEMA_VERSION)).toBe(1);

    // No command could author it and the save leg's fallback would not have
    // minted from it either, so there is no section to keep.
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

  test("every version below the current one still carries exactly one step", () => {
    expect(attrSchemaMigrationSteps(LEGACY_ATTR_SCHEMA_VERSION)).toHaveLength(1);
    expect(attrSchemaMigrationSteps(0)).toHaveLength(FOLIO_YJS_ATTR_SCHEMA_VERSION);
  });
});
