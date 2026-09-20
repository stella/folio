import { Result, TaggedError, panic } from "better-result";
import type { Node as PMNode } from "prosemirror-model";
import type * as Y from "yjs";

import {
  PROSE_PARAGRAPH_SOURCE_CONTRACT_ATTR,
  getProseDocumentParagraphPropertySourceContract,
} from "../docx/paragraphPropertySource";

/**
 * Folio's own keys inside a collaboration document: one map, two independent
 * bindings.
 *
 * `paragraphSourceContract` binds the snapshot to the exact bytes of the DOCX
 * it was seeded from. `attrSchemaVersion` binds it to the shape of the node
 * attrs Folio persists. A snapshot can be current on one and stale on the
 * other, so neither key may be derived from the other.
 */
const FOLIO_YJS_METADATA_MAP_NAME = "folio:document-metadata";
const PARAGRAPH_SOURCE_CONTRACT_KEY = "paragraphSourceContract";
const ATTR_SCHEMA_VERSION_KEY = "attrSchemaVersion";

/**
 * Carry a snapshot written under `version` forward to the next version,
 * rewriting the fragment in place and answering how many paragraphs changed.
 */
type AttrSchemaMigrationStep = (fragment: Y.XmlFragment) => number;

/**
 * Version 0 is every snapshot written before the marker existed. The marker
 * shipped without changing a single attr, so carrying a v0 snapshot to v1 is
 * the identity: it only stamps the key.
 */
const stampMarkerOnly: AttrSchemaMigrationStep = () => 0;

/** The node types whose `fldLock` and `dirty` attrs version 2 rewrites. */
const FIELD_ELEMENT_NAMES = new Set(["field", "structuredField"]);
const STATED_FLAG_ATTRS = ["fldLock", "dirty"] as const;

/**
 * Version 1 stored `fldLock` and `dirty` as booleans defaulting to `false`,
 * because the reader that filled them tested `=== true`. A field that authored
 * an explicit `w:fldLock="0"` and one that authored nothing were both stored as
 * `false`, so `false` never meant an explicit off and cannot be kept as one now
 * that `null` is the absence: a v1 snapshot's `false` would start writing an
 * attribute the document never carried.
 */
const dropUnstatedFieldFlags: AttrSchemaMigrationStep = (fragment) => {
  let rewritten = 0;
  const visit = (node: Y.XmlElement | Y.XmlFragment): void => {
    if ("nodeName" in node && FIELD_ELEMENT_NAMES.has(node.nodeName)) {
      // A Yjs attribute holds JSON, not a string; the typings say otherwise.
      const attributes: Record<string, unknown> = node.getAttributes();
      let changed = false;
      for (const attr of STATED_FLAG_ATTRS) {
        if (attributes[attr] === false) {
          node.removeAttribute(attr);
          changed = true;
        }
      }
      if (changed) {
        rewritten += 1;
      }
    }
    for (const child of node.toArray()) {
      if (typeof child !== "string" && "toArray" in child) {
        visit(child);
      }
    }
  };
  visit(fragment);
  return rewritten;
};

/**
 * Version 3 adds `docxRotation`, `docxFlipH` and `docxFlipV` to the drawing
 * nodes. A v2 snapshot states none of them, and `readAuthoredTransform` reads
 * such a node from its `transform` CSS, which is the only record it ever had,
 * so nothing has to be rewritten. The marker still moves: a v3 snapshot read
 * by a v2 build would have the three attrs dropped without a trace.
 */
const drawingTransformAttrsAreAdditive: AttrSchemaMigrationStep = () => 0;

/** The node types whose stated cell width version 4 backfills. */
const TABLE_CELL_ELEMENT_NAMES = new Set(["tableCell", "tableHeader"]);

/**
 * Write a node attribute the binding stores as JSON.
 *
 * The counterpart of the `getAttributes` reads above: `y-prosemirror` keeps a
 * node's attrs as the values ProseMirror holds, while Yjs types `setAttribute`
 * for the XML use that the same class also serves.
 */
const setJsonAttribute = (node: Y.XmlElement, name: string, value: object): void => {
  // SAFETY: the value is read back by `y-prosemirror` as the object written
  // here; narrowing it to the `string` Yjs's XML typings name would be the
  // wrong shape rather than a safer one, as the reads above already show.
  node.setAttribute(name, value as unknown as string);
};

const statedWidthOf = (attributes: Record<string, unknown>): object | undefined => {
  const original = attributes["_originalFormatting"];
  if (typeof original !== "object" || original === null || !("width" in original)) {
    return undefined;
  }
  const { width } = original;
  if (typeof width !== "object" || width === null) {
    return undefined;
  }
  // The cell's current width, which a resize may have moved off the one the
  // source stated, under the type the source stated it in.
  const value = attributes["width"];
  return typeof value === "number"
    ? { value, type: attributes["widthType"] ?? Reflect.get(width, "type") }
    : width;
};

/**
 * Version 4 adds `_authoredWidth` to the table cells. Up to version 3 the save
 * leg wrote `w:tcW` from `width`, which carries the width the *table* resolved
 * for a cell that states none, so a v3 snapshot cannot be read as-is under a
 * save leg that writes only what the cell states: every authored `w:tcW` would
 * go. `_originalFormatting.width` is the record of which cells stated one, so
 * the step backfills from it, at the width the cell currently holds.
 *
 * A cell resized in a v3 snapshot that stated no width of its own keeps no
 * preferred width; its column geometry rides `w:tblGrid`, which the table's
 * `columnWidths` still carries.
 */
const backfillStatedCellWidths: AttrSchemaMigrationStep = (fragment) => {
  let rewritten = 0;
  const visit = (node: Y.XmlElement | Y.XmlFragment): void => {
    if ("nodeName" in node && TABLE_CELL_ELEMENT_NAMES.has(node.nodeName)) {
      // A Yjs attribute holds JSON, not a string; the typings say otherwise.
      const attributes: Record<string, unknown> = node.getAttributes();
      const stated = statedWidthOf(attributes);
      if (stated !== undefined) {
        setJsonAttribute(node, "_authoredWidth", stated);
        rewritten += 1;
      }
    }
    for (const child of node.toArray()) {
      if (typeof child !== "string" && "toArray" in child) {
        visit(child);
      }
    }
  };
  visit(fragment);
  return rewritten;
};

/** Every attr-schema version this build reads, oldest first, with no gaps. */
const FOLIO_YJS_ATTR_SCHEMA_VERSIONS = [0, 1, 2, 3, 4] as const;

/** An attr-schema version this build can read. */
export type FolioYjsAttrSchemaVersion = (typeof FOLIO_YJS_ATTR_SCHEMA_VERSIONS)[number];

/**
 * The one place a persisted attr-shape change is decided.
 *
 * Changing the shape of any node or mark attr Folio persists means: append the
 * new version to {@link FOLIO_YJS_ATTR_SCHEMA_VERSIONS}, map the previous
 * version to the step that rewrites it, and map the new version to `"current"`.
 * The map is total over the version union, so a version added without a
 * decision does not compile, and `migrateFolioYjsSnapshot` picks the new step
 * up without being touched.
 */
const ATTR_SCHEMA_MIGRATIONS = {
  0: stampMarkerOnly,
  1: dropUnstatedFieldFlags,
  2: drawingTransformAttrsAreAdditive,
  3: backfillStatedCellWidths,
  4: "current",
} as const satisfies Record<FolioYjsAttrSchemaVersion, AttrSchemaMigrationStep | "current">;

type CurrentAttrSchemaVersion = {
  [Version in FolioYjsAttrSchemaVersion]: (typeof ATTR_SCHEMA_MIGRATIONS)[Version] extends "current"
    ? Version
    : never;
}[FolioYjsAttrSchemaVersion];

/**
 * The attr-schema version this build writes. Derived against the migration map
 * so the constant and the map cannot disagree.
 */
export const FOLIO_YJS_ATTR_SCHEMA_VERSION = 4 satisfies CurrentAttrSchemaVersion;

/**
 * The steps that carry a snapshot written under `fromVersion` up to
 * {@link FOLIO_YJS_ATTR_SCHEMA_VERSION}, in order. Empty when the snapshot is
 * already current.
 */
export const attrSchemaMigrationSteps = (
  fromVersion: FolioYjsAttrSchemaVersion,
): readonly AttrSchemaMigrationStep[] =>
  FOLIO_YJS_ATTR_SCHEMA_VERSIONS.filter((version) => version >= fromVersion).flatMap((version) => {
    const step = ATTR_SCHEMA_MIGRATIONS[version];
    return step === "current" ? [] : [step];
  });

/**
 * Carry a fragment this build is about to read or write up to the attr shape
 * this build writes, and stamp the marker.
 *
 * Every entry point that hands a fragment to `initProseMirrorDoc` calls this,
 * for two reasons that are the same reason. A step that rewrites values has to
 * run before a node is built from them, or the build reads the old shape as
 * the new one. And the marker has to say what the fragment may now hold before
 * the first edit writes an attr of this build's shape into it, because an
 * older build reading an unmarked snapshot drops what it does not know without
 * a trace.
 *
 * Returns how many elements the steps rewrote.
 */
export const applyAttrSchemaMigrations = (
  ydoc: Y.Doc,
  fragment: Y.XmlFragment,
  fromVersion: FolioYjsAttrSchemaVersion,
): number => {
  if (fromVersion === FOLIO_YJS_ATTR_SCHEMA_VERSION) {
    return 0;
  }
  let rewritten = 0;
  ydoc.transact(() => {
    for (const step of attrSchemaMigrationSteps(fromVersion)) {
      rewritten += step(fragment);
    }
    writeYjsAttrSchemaVersion(ydoc);
  });
  return rewritten;
};

/** Raised when a snapshot's attr-schema marker is ahead of the running code. */
export class FolioYjsAttrSchemaVersionError extends TaggedError("FolioYjsAttrSchemaVersionError")<{
  message: string;
  /** The marker exactly as stored, so the host can log what it refused. */
  marker: unknown;
  /** The newest attr-schema version this build understands. */
  supportedVersion: FolioYjsAttrSchemaVersion;
}> {}

const isReadableVersion = (marker: unknown): marker is FolioYjsAttrSchemaVersion =>
  typeof marker === "number" &&
  Number.isInteger(marker) &&
  marker >= 0 &&
  marker <= FOLIO_YJS_ATTR_SCHEMA_VERSION;

/**
 * The attr-schema version a collaboration document was written under.
 *
 * An absent marker is version 0: every snapshot predates the marker, so no
 * backfill is needed to classify one. An older version loads, because
 * {@link attrSchemaMigrationSteps} says what it means. A marker this
 * build does not know is a snapshot written by newer code, and reading it would
 * be a silent misread: ProseMirror copies unknown attr values into the node
 * verbatim and drops unknown keys without a trace. It fails instead.
 */
export const readYjsAttrSchemaVersion = (
  ydoc: Y.Doc,
): Result<FolioYjsAttrSchemaVersion, FolioYjsAttrSchemaVersionError> => {
  const marker: unknown = ydoc.getMap(FOLIO_YJS_METADATA_MAP_NAME).get(ATTR_SCHEMA_VERSION_KEY);
  if (marker === undefined) {
    return Result.ok(0);
  }
  if (isReadableVersion(marker)) {
    return Result.ok(marker);
  }
  return Result.err(
    new FolioYjsAttrSchemaVersionError({
      marker,
      message:
        "The collaboration snapshot was written by a newer Folio attr schema than this build reads.",
      supportedVersion: FOLIO_YJS_ATTR_SCHEMA_VERSION,
    }),
  );
};

export const proseDocumentParagraphSourceContract = (document: PMNode): string | null => {
  return getProseDocumentParagraphPropertySourceContract(document);
};

/** Stamp the attr-schema version of the build that wrote the current fragment. */
export const writeYjsAttrSchemaVersion = (ydoc: Y.Doc): void => {
  ydoc
    .getMap(FOLIO_YJS_METADATA_MAP_NAME)
    .set(ATTR_SCHEMA_VERSION_KEY, FOLIO_YJS_ATTR_SCHEMA_VERSION);
};

/**
 * Write both metadata bindings for a freshly seeded collaboration document.
 * The attr-schema version is written unconditionally; the paragraph-source
 * contract only when the seeding document carries one, because a document
 * built in memory rather than parsed from a package has no source to bind to.
 */
export const writeYjsDocumentMetadata = (ydoc: Y.Doc, document: PMNode): void => {
  writeYjsAttrSchemaVersion(ydoc);
  const contract = proseDocumentParagraphSourceContract(document);
  if (contract) {
    ydoc.getMap(FOLIO_YJS_METADATA_MAP_NAME).set(PARAGRAPH_SOURCE_CONTRACT_KEY, contract);
  }
};

export const readYjsParagraphSourceContract = (ydoc: Y.Doc): string | null => {
  const contract = ydoc.getMap(FOLIO_YJS_METADATA_MAP_NAME).get(PARAGRAPH_SOURCE_CONTRACT_KEY);
  return typeof contract === "string" ? contract : null;
};

export const withParagraphSourceContract = (document: PMNode, contract: string): PMNode => {
  if (document.type.name !== "doc") {
    panic("A paragraph-property source contract can only attach to a document node");
  }
  return document.type.create(
    { ...document.attrs, [PROSE_PARAGRAPH_SOURCE_CONTRACT_ATTR]: contract },
    document.content,
    document.marks,
  );
};
