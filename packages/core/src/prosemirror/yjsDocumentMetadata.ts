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

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null;

/**
 * A shared text node rather than an element.
 *
 * Told apart structurally so this module keeps its type-only import of Yjs: an
 * `instanceof` check would make it a value import, and the class identity would
 * then have to match the Yjs instance the host loaded.
 */
const isXmlText = (child: object): child is Y.XmlText => "toDelta" in child;

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

/**
 * Version 4 renames the `pageBreakRunOwner` mark to `runIdentity`, which is a
 * value rewrite rather than an additive change, and the only mark-attr change
 * so far.
 *
 * Mark attrs persist as Y.Text delta attributes keyed by the mark's name, and
 * `createTextNodesFromYText` rebuilds each one with `schema.mark(name, attrs)`
 * inside a `try` whose `catch` **deletes the Y.Text item** and persists the
 * deletion. A v3 snapshot holding `pageBreakRunOwner` read by this build would
 * therefore lose the text under the mark, not merely the mark. Every step
 * before this one walked `Y.XmlElement` attributes; this walks the delta of
 * every `Y.XmlText` and re-formats each range that states the old name.
 *
 * The payload is not backfilled. A snapshot has no access to the DOCX it was
 * seeded from, so an old room keeps today's behaviour — an identity with no
 * remainder — until it is reseeded, and a new room gets the remainder.
 */
const renamePageBreakRunOwnerMarkAttr: AttrSchemaMigrationStep = (fragment) => {
  const OLD_MARK_NAME = "pageBreakRunOwner";
  const NEW_MARK_NAME = "runIdentity";
  let rewritten = 0;

  const rewriteText = (text: Y.XmlText): void => {
    let index = 0;
    // The delta is read whole before anything is formatted: `format` rewrites
    // the very structure being walked, and a range's offset is only valid
    // against the delta it came from.
    const ranges: { at: number; length: number; owner: unknown }[] = [];
    for (const op of text.toDelta()) {
      const insert: unknown = op.insert;
      const length = typeof insert === "string" ? insert.length : 1;
      const attributes: unknown = op.attributes;
      if (isRecord(attributes) && OLD_MARK_NAME in attributes) {
        ranges.push({ at: index, length, owner: attributes[OLD_MARK_NAME] });
      }
      index += length;
    }
    for (const { at, length, owner } of ranges) {
      // The old mark carried `{id}` and nothing else, so the new mark's
      // payload fields stay absent, which is what "this run carries no
      // remainder" spells.
      const identity = isRecord(owner) && typeof owner["id"] === "number" ? { id: owner["id"] } : {};
      text.format(at, length, { [OLD_MARK_NAME]: null, [NEW_MARK_NAME]: identity });
      rewritten += 1;
    }
  };

  const visit = (node: Y.XmlElement | Y.XmlFragment): void => {
    for (const child of node.toArray()) {
      if (typeof child === "string") {
        continue;
      }
      if (isXmlText(child)) {
        rewriteText(child);
        continue;
      }
      if ("toArray" in child) {
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
  3: renamePageBreakRunOwnerMarkAttr,
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
