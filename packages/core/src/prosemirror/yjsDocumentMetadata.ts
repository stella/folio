import { Result, TaggedError, panic } from "better-result";
import { DRAWING_RAW_XML_MODES, outlineLevelFromStatedValue } from "@stll/docx-core/model";
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

/** The node type whose `outlineLevel` attr version 7 rewrites. */
const PARAGRAPH_ELEMENT_NAME = "paragraph";

/**
 * Version 6 stored `outlineLevel` as the `w:outlineLvl w:val` number, with 9
 * meaning body text and every consumer deciding that for itself. Version 7
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

/** The node type whose relationship and raw-XML attrs version 4 rewrites. */
const IMAGE_ELEMENT_NAME = "image";

/**
 * Stands in for the capture marker a node migrated to `previewOnly` never had.
 *
 * `drawingFromImageAttrs` reads the marker's presence and never its value: the
 * editor projection is lossy, so a value compared against the image rebuilt
 * from it would report an untouched round trip as an edit. What the marker has
 * to say here is that the node still holds the capture it was projected with,
 * which is exactly the condition the step selects on.
 */
const CARRIED_PREVIEW_CAPTURE = "carriedPreviewCapture";

/** Rewrite one stored image node, answering whether anything changed. */
const migrateImageNode = (node: Y.XmlElement): boolean => {
  // A Yjs attribute holds JSON, not a string; the typings say otherwise.
  const attributes: Record<string, unknown> = node.getAttributes();
  const rId = attributes["rId"];
  const dropsEmptyId = rId === "";
  const classifies =
    !(typeof rId === "string" && rId.length > 0) &&
    typeof attributes["_docxRawXml"] === "string" &&
    attributes["_docxRawXmlMode"] === undefined;
  if (dropsEmptyId) {
    node.removeAttribute("rId");
  }
  if (classifies) {
    node.setAttribute("_docxRawXmlMode", DRAWING_RAW_XML_MODES.PREVIEW_ONLY);
    node.setAttribute("_docxRawImageFingerprint", CARRIED_PREVIEW_CAPTURE);
  }
  return dropsEmptyId || classifies;
};

/**
 * Version 3 spelled "this drawing carries no relationship" two ways.
 *
 * `rId: ""` is not a key any relationship answers to, so an image that stated
 * it was a drawing with no relationship saying it had one: it reached a save
 * as `r:embed=""`, and only the lookups that knew to ask about both spellings
 * read it as the absence it was. The model has one spelling now, so the attr
 * loses the other.
 *
 * The same pass classifies the drawings that spelling hid. A VML shape folio
 * cannot project is carried as a render of its markup, and version 3 projected
 * that render as an ordinary editable picture: an edit dropped the capture, and
 * the save wrote the render into the package as a picture in place of the
 * shape. An image node that holds a capture, names no relationship and states
 * no mode is exactly that drawing, because every other captured drawing either
 * names a relationship or states its mode. Still holding the capture is what
 * says the node has not been edited, which is what the stamped marker means.
 */
const classifyUnrelatedCapturedDrawings: AttrSchemaMigrationStep = (fragment) => {
  let rewritten = 0;
  const visit = (node: Y.XmlElement | Y.XmlFragment): void => {
    if ("nodeName" in node && node.nodeName === IMAGE_ELEMENT_NAME && migrateImageNode(node)) {
      rewritten += 1;
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

/** The node types whose stated cell width version 5 backfills. */
const TABLE_CELL_ELEMENT_NAMES = new Set(["tableCell", "tableHeader"]);

/** Write a node attribute the binding stores as JSON. */
const setJsonAttribute = (node: Y.XmlElement, name: string, value: object): void => {
  // SAFETY: y-prosemirror reads this value back as the object written here;
  // Yjs's XML typing names a string even though the binding stores JSON.
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
  const value = attributes["width"];
  return typeof value === "number"
    ? { value, type: attributes["widthType"] ?? Reflect.get(width, "type") }
    : width;
};

/**
 * Version 5 adds `_authoredWidth` to table cells. Version 4 wrote `w:tcW`
 * from the resolved cell width, including cells that stated no preferred
 * width. `_originalFormatting.width` records which cells authored one, so
 * only those cells receive the new attr.
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

/**
 * Version 6 gives `tableRow`'s `hidden` attr the `null` default the other
 * tri-states carry. Up to version 5 it defaulted to `false`, and the reader
 * that filled it took only an explicit on, so a row that authored
 * `<w:hidden w:val="0"/>` and one that authored nothing were both stored as
 * `false`. `false` therefore never meant an explicit off and cannot be kept as
 * one: a version-5 snapshot's `false` would start writing an element the
 * document never carried. The explicit offs it does hold travelled as captured
 * bytes on `_originalFormatting`, which the step leaves alone.
 */
const dropUnstatedRowHidden: AttrSchemaMigrationStep = (fragment) => {
  let rewritten = 0;
  const visit = (node: Y.XmlElement | Y.XmlFragment): void => {
    if ("nodeName" in node && node.nodeName === "tableRow") {
      // A Yjs attribute holds JSON, not a string; the typings say otherwise.
      const attributes: Record<string, unknown> = node.getAttributes();
      if (attributes["hidden"] === false) {
        node.removeAttribute("hidden");
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
const FOLIO_YJS_ATTR_SCHEMA_VERSIONS = [0, 1, 2, 3, 4, 5, 6, 7] as const;

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
  3: classifyUnrelatedCapturedDrawings,
  4: backfillStatedCellWidths,
  5: dropUnstatedRowHidden,
  6: outlineLevelBecomesAUnion,
  7: "current",
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
export const FOLIO_YJS_ATTR_SCHEMA_VERSION = 7 satisfies CurrentAttrSchemaVersion;

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
