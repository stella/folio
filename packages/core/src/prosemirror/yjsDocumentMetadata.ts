import { Result, TaggedError, panic } from "better-result";
import {
  DRAWING_RAW_XML_MODES,
  outlineLevelFromStatedValue,
  type OutlineLevel,
  type ParagraphNumberingOverride,
  paragraphNumberingFromSlots,
} from "@stll/docx-core/model";
import type { Node as PMNode } from "prosemirror-model";
import type * as Y from "yjs";

import {
  PROSE_PARAGRAPH_SOURCE_CONTRACT_ATTR,
  getProseDocumentParagraphPropertySourceContract,
} from "../docx/paragraphPropertySource";
import { mintSectionProperties, parseSectionBreakType } from "./sectionCarrier";

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

/** A shared text node, identified structurally to keep Yjs a type-only import. */
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

/** The node type whose `outlineLevel` attrs version 7 rewrites. */
const PARAGRAPH_ELEMENT_NAME = "paragraph";

/** A stored outline level, or the absence of one in this carrier. */
const versionSixOutlineLevel = (value: unknown): OutlineLevel | null | undefined => {
  if (typeof value !== "number") {
    return undefined;
  }
  return outlineLevelFromStatedValue(value) ?? null;
};

/** Rewrite one persisted paragraph-formatting carrier when it stores a v6 number. */
const versionSixFormatting = (value: unknown): Record<string, unknown> | null => {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    return null;
  }
  const outlineLevel = versionSixOutlineLevel(Reflect.get(value, "outlineLevel"));
  if (outlineLevel === undefined) {
    return null;
  }
  const migrated: Record<string, unknown> = { ...value };
  if (outlineLevel === null) {
    Reflect.deleteProperty(migrated, "outlineLevel");
  } else {
    migrated["outlineLevel"] = outlineLevel;
  }
  return migrated;
};

/** Rewrite outline levels in paragraph-property-change formatting carriers. */
const versionSixPropertyChanges = (value: unknown): Record<string, unknown>[] | null => {
  if (!Array.isArray(value)) {
    return null;
  }
  const changes: Record<string, unknown>[] = [];
  let changed = false;
  for (const change of value) {
    if (typeof change !== "object" || change === null || Array.isArray(change)) {
      return null;
    }
    const migrated: Record<string, unknown> = { ...change };
    for (const tier of ["previousFormatting", "currentFormatting"] as const) {
      const formatting = versionSixFormatting(change[tier]);
      if (formatting !== null) {
        migrated[tier] = formatting;
        changed = true;
      }
    }
    changes.push(migrated);
  }
  return changed ? changes : null;
};

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
      let changed = false;
      const outlineLevel = versionSixOutlineLevel(attributes["outlineLevel"]);
      if (outlineLevel !== undefined) {
        if (outlineLevel === null) {
          node.removeAttribute("outlineLevel");
        } else {
          // @ts-expect-error — a Yjs attribute holds JSON; the typings narrow
          // to string, and the union is what this step exists to store.
          node.setAttribute("outlineLevel", outlineLevel);
        }
        changed = true;
      }
      for (const attr of ["_originalFormatting", "_resolvedFormatting"] as const) {
        const formatting = versionSixFormatting(attributes[attr]);
        if (formatting !== null) {
          // @ts-expect-error — a Yjs attribute holds JSON; the typings narrow
          // to string, and the migrated record is what this step stores.
          node.setAttribute(attr, formatting);
          changed = true;
        }
      }
      const propertyChanges = versionSixPropertyChanges(attributes["_propertyChanges"]);
      if (propertyChanges !== null) {
        // @ts-expect-error — as above.
        node.setAttribute("_propertyChanges", propertyChanges);
        changed = true;
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

/**
 * The `_propertyChanges` entries a paragraph carries, as a version-7 snapshot
 * stored them: an array of records whose `previousFormatting` may carry the
 * numbering the paragraph had before a `w:pPrChange`.
 */
const versionSevenPropertyChanges = (value: unknown): Record<string, unknown>[] | null => {
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
 * A value that already carries a `kind` is one of those. Version 7 wrote the
 * model's union into `_originalFormatting` and into a recorded
 * `currentFormatting` while the `numPr` attr beside them still held the two
 * slots, so a version-7 paragraph can hold both spellings at once: that
 * divergence is what version 8 exists to end.
 */
const versionSevenNumbering = (value: unknown): ParagraphNumberingOverride | null | undefined => {
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
const versionSevenFormatting = (value: unknown): Record<string, unknown> | null => {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    return null;
  }
  const stored: Record<string, unknown> = { ...value };
  let changed = false;
  for (const key of NUMBERING_ATTR_KEYS) {
    const migrated = versionSevenNumbering(stored[key]);
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
 * Version 7 stored `numPr` as the two `<w:numPr>` slots, with the reserved
 * `numId` 0 for a cancellation and a bare `ilvl` for a level stated without an
 * id. Version 8 stores the model's union, so the pair has to be mapped:
 * `numId` 0 becomes `none` whatever level sat beside it (a cancellation names
 * no id for a level to belong to), an `ilvl` without a `numId` becomes
 * `levelOnly`, the two together become `reference`, and an object stating
 * neither slot stated nothing is dropped. An absent attr stays absent.
 *
 * Like the version-7 step it cannot be left to a lazy read: ProseMirror copies
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
        const migrated = versionSevenNumbering(attributes[attr]);
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
      const originalFormatting = versionSevenFormatting(attributes["_originalFormatting"]);
      if (originalFormatting !== null) {
        changed = true;
        // @ts-expect-error — as above.
        node.setAttribute("_originalFormatting", originalFormatting);
      }
      const changes = versionSevenPropertyChanges(attributes["_propertyChanges"]);
      if (changes !== null) {
        let changesChanged = false;
        const migratedChanges: Record<string, unknown>[] = [];
        for (const change of changes) {
          const migrated = Object.assign({}, change);
          for (const tier of ["previousFormatting", "currentFormatting"] as const) {
            const formatting = versionSevenFormatting(change[tier]);
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

/**
 * Version 9 makes `_sectionProperties` the one carrier of a section break.
 * Earlier snapshots could state `sectionBreakType` instead, and the save leg
 * minted a `SectionProperties` from it. The attr is gone from the schema, so
 * the migration must mint that record before ProseMirror drops the old key.
 * A paragraph that already has the record keeps it as the authority.
 */
const SECTION_BREAK_TYPE_ATTR = "sectionBreakType";
const SECTION_PROPERTIES_ATTR = "_sectionProperties";

const mintSectionPropertiesFromBreakType: AttrSchemaMigrationStep = (fragment) => {
  let rewritten = 0;
  const visit = (node: Y.XmlElement | Y.XmlFragment): void => {
    if ("nodeName" in node && node.nodeName === PARAGRAPH_ELEMENT_NAME) {
      // A Yjs attribute holds JSON, not a string; the typings say otherwise.
      const attributes: Record<string, unknown> = node.getAttributes();
      if (SECTION_BREAK_TYPE_ATTR in attributes) {
        const breakType = parseSectionBreakType(attributes[SECTION_BREAK_TYPE_ATTR]);
        if (breakType !== null && attributes[SECTION_PROPERTIES_ATTR] == null) {
          // SAFETY: a Yjs attribute holds JSON and is read back as the attr's
          // own shape; only `Y.XmlElement`'s typings narrow it to `string`.
          node.setAttribute(
            SECTION_PROPERTIES_ATTR,
            mintSectionProperties(breakType) as unknown as string,
          );
        }
        node.removeAttribute(SECTION_BREAK_TYPE_ATTR);
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

/** The node types whose `showingPlaceholder` attr version 10 rewrites. */
const SDT_ELEMENT_NAMES = new Set(["sdt", "blockSdt"]);

/**
 * Version 9 stored `showingPlaceholder` as a boolean defaulting to `false`.
 * A control that authored an explicit off and one that authored nothing were
 * both stored as `false`, so the old false cannot become a newly authored
 * `<w:showingPlcHdr w:val="0"/>` when the absence is now `null`.
 */
const dropUnstatedPlaceholderFlags: AttrSchemaMigrationStep = (fragment) => {
  let rewritten = 0;
  const visit = (node: Y.XmlElement | Y.XmlFragment): void => {
    if ("nodeName" in node && SDT_ELEMENT_NAMES.has(node.nodeName)) {
      const attributes: Record<string, unknown> = node.getAttributes();
      if (attributes["showingPlaceholder"] === false) {
        node.removeAttribute("showingPlaceholder");
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
 * Version 11 renames the persisted `pageBreakRunOwner` mark to `runIdentity`.
 * Unknown mark names make y-prosemirror delete their covered text, so the
 * Y.Text delta must be rewritten before the ProseMirror document is built.
 */
const renamePageBreakRunOwnerMarkAttr: AttrSchemaMigrationStep = (fragment) => {
  const OLD_MARK_NAME = "pageBreakRunOwner";
  const NEW_MARK_NAME = "runIdentity";
  let rewritten = 0;

  const rewriteText = (sharedText: Y.XmlText): void => {
    let index = 0;
    const ranges: { at: number; length: number; owner: unknown }[] = [];
    for (const op of sharedText.toDelta()) {
      const insert: unknown = op.insert;
      const length = typeof insert === "string" ? insert.length : 1;
      const attributes: unknown = op.attributes;
      if (isRecord(attributes) && OLD_MARK_NAME in attributes) {
        ranges.push({ at: index, length, owner: attributes[OLD_MARK_NAME] });
      }
      index += length;
    }
    for (const { at, length, owner } of ranges) {
      const identity = isRecord(owner) && typeof owner["id"] === "number" ? { id: owner["id"] } : {};
      sharedText.format(at, length, { [OLD_MARK_NAME]: null, [NEW_MARK_NAME]: identity });
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
const FOLIO_YJS_ATTR_SCHEMA_VERSIONS = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11] as const;

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
  7: numberingBecomesAUnion,
  8: mintSectionPropertiesFromBreakType,
  9: dropUnstatedPlaceholderFlags,
  10: renamePageBreakRunOwnerMarkAttr,
  11: "current",
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
export const FOLIO_YJS_ATTR_SCHEMA_VERSION = 11 satisfies CurrentAttrSchemaVersion;

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
