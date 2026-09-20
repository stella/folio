/**
 * The snapshot migrator is the only offline path that rewrites persisted
 * editor state, and a host runs it over a whole corpus in one pass. Two things
 * have to hold over arbitrary snapshots rather than over one fixture: it must
 * be a fixed point (a corpus can be swept twice, a room can be swept after a
 * partial failure), and it must not touch content — the v0 step stamps the
 * marker and nothing else.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import { initProseMirrorDoc, prosemirrorToYXmlFragment } from "y-prosemirror";
import * as Y from "yjs";

import {
  type OutlineLevel,
  outlineLevelFromStatedValue,
  type ParagraphNumberingOverride,
  paragraphNumberingFromSlots,
} from "@stll/docx-core/model";

import { propertyConfig, propertyTestTimeout } from "../../../../../test/property-testing";
import { schema } from "../../prosemirror/schema";
import {
  FOLIO_YJS_ATTR_SCHEMA_VERSION,
  writeYjsDocumentMetadata,
} from "../../prosemirror/yjsDocumentMetadata";
import { FOLIO_YJS_PROSEMIRROR_FRAGMENT_NAME } from "./materializeYjsDocx";
import { migrateFolioYjsSnapshot } from "./migrateYjsAttrSchema";

/**
 * Paragraphs carrying the attrs whose shape the union migrations change. Both
 * are generated in the shape the older snapshot stored — `outlineLevel` as the
 * `w:outlineLvl w:val` number, `numPr` as the two `<w:numPr>` slots — because
 * that is what the steps have to read.
 */
const paragraphs = fc.array(
  fc.record({
    numPr: fc.option(
      fc.record({ ilvl: fc.nat({ max: 8 }), numId: fc.nat({ max: 9 }) }, { requiredKeys: [] }),
      { nil: null },
    ),
    outlineLevel: fc.option(fc.nat({ max: 12 }), { nil: null }),
    text: fc.string({ maxLength: 24, minLength: 0 }),
  }),
  { maxLength: 6, minLength: 1 },
);

type SyntheticParagraph = typeof paragraphs extends fc.Arbitrary<infer T> ? T[number] : never;

/** The union a stored version-6 number carries forward to, or absent. */
const migratedOutlineLevel = (stated: number | null): OutlineLevel | null =>
  stated === null ? null : (outlineLevelFromStatedValue(stated) ?? null);

/** The union a stored version-7 slot pair carries forward to, or absent. */
const migratedNumbering = (
  stated: SyntheticParagraph["numPr"],
): ParagraphNumberingOverride | null =>
  stated === null ? null : (paragraphNumberingFromSlots(stated) ?? null);

type StoredAttrsOf = (block: SyntheticParagraph) => { numPr: unknown; outlineLevel: unknown };

const proseDocument = (blocks: readonly SyntheticParagraph[], attrsOf: StoredAttrsOf) =>
  schema.topNodeType.create(
    null,
    blocks.map((block) =>
      // Rebuilt as a plain object: Yjs refuses an attribute value whose
      // constructor is not `Object`, and fast-check's records are
      // null-prototype.
      schema.nodes["paragraph"]?.create(
        attrsOf(block),
        block.text.length === 0 ? null : schema.text(block.text),
      ),
    ),
  );

/** The attrs a pre-marker build stored: both fields in their old shapes. */
const legacyAttrs: StoredAttrsOf = ({ numPr, outlineLevel }) => ({
  numPr: numPr === null ? null : { ...numPr },
  outlineLevel,
});

/** The attrs this build stores, which is what a swept snapshot must equal. */
const currentAttrs: StoredAttrsOf = ({ numPr, outlineLevel }) => ({
  numPr: migratedNumbering(numPr),
  outlineLevel: migratedOutlineLevel(outlineLevel),
});

type SyntheticSnapshot = {
  /** The update a pre-marker build would have stored. */
  unversioned: Uint8Array;
  /** The same content with this build's marker. */
  versioned: Uint8Array;
};

const syntheticSnapshot = (blocks: readonly SyntheticParagraph[]): SyntheticSnapshot => {
  const document = proseDocument(blocks, legacyAttrs);
  const unversionedDoc = new Y.Doc();
  prosemirrorToYXmlFragment(
    document,
    unversionedDoc.getXmlFragment(FOLIO_YJS_PROSEMIRROR_FRAGMENT_NAME),
  );
  const unversioned = Y.encodeStateAsUpdate(unversionedDoc);
  unversionedDoc.destroy();

  const versionedDoc = new Y.Doc();
  Y.applyUpdate(versionedDoc, unversioned);
  writeYjsDocumentMetadata(versionedDoc, document);
  const versioned = Y.encodeStateAsUpdate(versionedDoc);
  versionedDoc.destroy();

  return { unversioned, versioned };
};

const proseDocumentOf = (update: Uint8Array) => {
  const ydoc = new Y.Doc();
  Y.applyUpdate(ydoc, update);
  const document = initProseMirrorDoc(
    ydoc.getXmlFragment(FOLIO_YJS_PROSEMIRROR_FRAGMENT_NAME),
    schema,
  ).doc;
  ydoc.destroy();
  return document;
};

const expectMigrated = (update: Uint8Array) => {
  const migrated = migrateFolioYjsSnapshot(update);
  if (migrated.isErr()) {
    throw migrated.error;
  }
  return migrated.value;
};

describe("migrateFolioYjsSnapshot", () => {
  test(
    "is a fixed point and leaves content alone",
    () => {
      fc.assert(
        fc.property(paragraphs, (blocks) => {
          const { unversioned } = syntheticSnapshot(blocks);
          const once = expectMigrated(unversioned);
          const twice = expectMigrated(once.update);

          expect(once.fromVersion).toBe(0);
          expect(once.toVersion).toBe(FOLIO_YJS_ATTR_SCHEMA_VERSION);
          // Each step counts the paragraphs it rewrote, and a paragraph
          // stating both attrs is rewritten by both.
          expect(once.paragraphsRewritten).toBe(
            blocks.filter(({ outlineLevel }) => outlineLevel !== null).length +
              blocks.filter(({ numPr }) => numPr !== null).length,
          );
          expect(twice.fromVersion).toBe(FOLIO_YJS_ATTR_SCHEMA_VERSION);
          expect(twice.update).toEqual(once.update);
          // Every attr equals what the stored value implies: the outline union
          // for 0..9 and absent beyond it, the numbering union for the slot
          // pair and absent when the element stated neither slot.
          expect(proseDocumentOf(once.update).eq(proseDocument(blocks, currentAttrs))).toBe(true);
        }),
        propertyConfig({ numRuns: 60 }),
      );
    },
    propertyTestTimeout(20_000),
  );

  test(
    "returns a current snapshot byte for byte",
    () => {
      fc.assert(
        fc.property(paragraphs, (blocks) => {
          const { versioned } = syntheticSnapshot(blocks);
          const migrated = expectMigrated(versioned);

          expect(migrated.fromVersion).toBe(FOLIO_YJS_ATTR_SCHEMA_VERSION);
          expect(migrated.update).toBe(versioned);
        }),
        propertyConfig({ numRuns: 60 }),
      );
    },
    propertyTestTimeout(20_000),
  );
});
