import { Result, TaggedError, panic } from "better-result";
import {
  outlineLevelFromStatedValue,
  type ParagraphNumberingOverride,
  paragraphNumberingFromSlots,
} from "@stll/docx-core/model";
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

/** The node type whose `outlineLevel` attr version 4 rewrites. */
const PARAGRAPH_ELEMENT_NAME = "paragraph";

/**
 * Version 3 stored `outlineLevel` as the `w:outlineLvl w:val` number, with 9
 * meaning body text and every consumer deciding that for itself. Version 4
 * stores `OutlineLevel`, so the number has to be mapped: 0..8 become the
 * heading arm, 9 becomes the body-text arm, and anything else is dropped,
 * matching the parse boundary. It cannot be left alone and read lazily,
 * because ProseMirror copies a stored attr into the node without validating
 * and the strict validator would then panic on the first read of an untouched
 * room.
 */
const outlineLevelBecomesAUnion: AttrSchemaMigrationStep = (fragment) => {
  let rewritten = 0;
  const visit = (node: Y.XmlElement | Y.XmlFragment): void => {
    if ("nodeName" in node && node.nodeName === PARAGRAPH_ELEMENT_NAME) {
      // A Yjs attribute holds JSON, not a string; the typings say otherwise.
      const attributes: Record<string, unknown> = node.getAttributes();
      const stated = attributes["outlineLevel"];
      if (typeof stated === "number") {
        const outlineLevel = outlineLevelFromStatedValue(stated);
        if (outlineLevel === undefined) {
          node.removeAttribute("outlineLevel");
        } else {
          // @ts-expect-error — a Yjs attribute holds JSON; the typings narrow
          // to string, and the union is what this step exists to store.
          node.setAttribute("outlineLevel", outlineLevel);
        }
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
 * The `_propertyChanges` entries a paragraph carries, as a version-4 snapshot
 * stored them: an array of records whose `previousFormatting` may carry the
 * numbering the paragraph had before a `w:pPrChange`.
 */
const versionFourPropertyChanges = (value: unknown): Record<string, unknown>[] | null => {
  if (!Array.isArray(value)) {
    return null;
  }
  const changes: Record<string, unknown>[] = [];
  for (const change of value) {
    if (typeof change !== "object" || change === null || Array.isArray(change)) {
      return null;
    }
    changes.push(change as Record<string, unknown>);
  }
  return changes;
};

/** The two paragraph attrs, and the two formatting fields, that state numbering. */
const NUMBERING_ATTR_KEYS = ["numPr", "numPrFromStyle"] as const;

/**
 * One stored two-slot value, as the union. `null` is an object that stated
 * neither slot, which stated nothing; `undefined` is a value this step must
 * leave exactly as it found it.
 *
 * A value that already carries a `kind` is one of those. Version 4 wrote the
 * model's union into `_originalFormatting` and into a recorded
 * `currentFormatting` while the `numPr` attr beside them still held the two
 * slots, so a version-4 paragraph can hold both spellings at once — that
 * divergence is what version 5 exists to end.
 */
const versionFourNumbering = (value: unknown): ParagraphNumberingOverride | null | undefined => {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    return undefined;
  }
  if (typeof Reflect.get(value, "kind") === "string") {
    return undefined;
  }
  const numId: unknown = Reflect.get(value, "numId");
  const ilvl: unknown = Reflect.get(value, "ilvl");
  return (
    paragraphNumberingFromSlots({
      ilvl: typeof ilvl === "number" ? ilvl : undefined,
      numId: typeof numId === "number" ? numId : undefined,
    }) ?? null
  );
};

/**
 * A stored formatting record with its numbering keys carried forward, or
 * `null` when it holds nothing this step rewrites. `null` inside the record
 * survives: in `previousFormatting` it is the tombstone for "the paragraph
 * carried no numbering before the change".
 */
const versionFourFormatting = (value: unknown): Record<string, unknown> | null => {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    return null;
  }
  const stored: Record<string, unknown> = { ...value };
  let changed = false;
  for (const key of NUMBERING_ATTR_KEYS) {
    const migrated = versionFourNumbering(stored[key]);
    if (migrated === undefined) {
      continue;
    }
    changed = true;
    if (migrated === null) {
      Reflect.deleteProperty(stored, key);
    } else {
      stored[key] = migrated;
    }
  }
  return changed ? stored : null;
};

/**
 * Version 4 stored `numPr` as the two `<w:numPr>` slots, with the reserved
 * `numId` 0 for a cancellation and a bare `ilvl` for a level stated without an
 * id. Version 5 stores the model's union, so the pair has to be mapped:
 * `numId` 0 becomes `none` whatever level sat beside it (a cancellation names
 * no id for a level to belong to), an `ilvl` without a `numId` becomes
 * `levelOnly`, the two together become `reference`, and an object stating
 * neither slot stated nothing is dropped. An absent attr stays absent.
 *
 * Like the version-4 step it cannot be left to a lazy read: ProseMirror copies
 * a stored attr into the node without validating, and
 * `optionalParagraphNumbering` would then refuse the first read of an
 * untouched room.
 */
const numberingBecomesAUnion: AttrSchemaMigrationStep = (fragment) => {
  let rewritten = 0;
  const visit = (node: Y.XmlElement | Y.XmlFragment): void => {
    if ("nodeName" in node && node.nodeName === PARAGRAPH_ELEMENT_NAME) {
      // A Yjs attribute holds JSON, not a string; the typings say otherwise.
      const attributes: Record<string, unknown> = node.getAttributes();
      let changed = false;
      for (const attr of NUMBERING_ATTR_KEYS) {
        const migrated = versionFourNumbering(attributes[attr]);
        if (migrated === undefined) {
          continue;
        }
        changed = true;
        if (migrated === null) {
          node.removeAttribute(attr);
        } else {
          // @ts-expect-error — a Yjs attribute holds JSON; the typings narrow
          // to string, and the union is what this step exists to store.
          node.setAttribute(attr, migrated);
        }
      }
      const originalFormatting = versionFourFormatting(attributes["_originalFormatting"]);
      if (originalFormatting !== null) {
        changed = true;
        // @ts-expect-error — as above.
        node.setAttribute("_originalFormatting", originalFormatting);
      }
      const changes = versionFourPropertyChanges(attributes["_propertyChanges"]);
      if (changes !== null) {
        let changesChanged = false;
        const migratedChanges: Record<string, unknown>[] = [];
        for (const change of changes) {
          const migrated = Object.assign({}, change);
          for (const tier of ["previousFormatting", "currentFormatting"] as const) {
            const formatting = versionFourFormatting(change[tier]);
            if (formatting !== null) {
              migrated[tier] = formatting;
              changesChanged = true;
            }
          }
          migratedChanges.push(migrated);
        }
        if (changesChanged) {
          changed = true;
          // @ts-expect-error — as above.
          node.setAttribute("_propertyChanges", migratedChanges);
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

/** Every attr-schema version this build reads, oldest first, with no gaps. */
const FOLIO_YJS_ATTR_SCHEMA_VERSIONS = [0, 1, 2, 3, 4, 5] as const;

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
  3: outlineLevelBecomesAUnion,
  4: numberingBecomesAUnion,
  5: "current",
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
export const FOLIO_YJS_ATTR_SCHEMA_VERSION = 5 satisfies CurrentAttrSchemaVersion;

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
