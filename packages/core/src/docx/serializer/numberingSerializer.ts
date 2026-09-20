/**
 * Numbering Serializer - Serialize numbering definitions back to OOXML XML
 *
 * Converts the parsed {@link NumberingDefinitions} model (abstract numberings
 * and concrete numbering instances) back into `word/numbering.xml`.
 * The inverse of `numberingParser`.
 *
 * OOXML Reference:
 * - Root: w:numbering
 * - Templates: w:abstractNum[@w:abstractNumId] with 0..8 w:lvl children
 * - Instances: w:num[@w:numId] referencing an abstractNum (+ optional overrides)
 *
 * Every container here writes its children in the order its content model
 * declares them, because that order is what the ordered sink's index means: a
 * capture recorded after the third field the reader kept goes back after the
 * third field this writes, so unmodelled markup keeps the neighbours it was
 * authored between.
 *
 * Two things the part carries are still not this writer's to rebuild. A custom
 * number format collapses to one of folio's synthetic pad-width names, and a
 * level's `w:pPr` and `w:rPr` keep only the subset their own readers model. So
 * the save paths keep splicing definition by definition rather than
 * overwriting the part: `rezip.ts` and `selectiveSave.ts` compare this output
 * against the same output taken over the original, which cancels both.
 */

import type {
  AbstractNumbering,
  LevelOverride,
  ListLevel,
  NumberFormat,
  NumberingDefinitions,
  NumberingInstance,
  ParagraphFormatting,
} from "../../types/document";
import { serializePreservedAttributes } from "../attributeRemainder";
import { serializeWithPreservedChildren } from "../containerChildren";
import { serializePartElement } from "./partNamespaces";
import { serializeTextFormatting } from "./textFormattingSerializer";
import { intAttr } from "./xmlUtils";
import { escapeXmlAttribute } from "@stll/docx-core";

/**
 * `w:numFmt`, as the model spells it.
 *
 * A level read from a `w:numFmt w:val="custom"` carries one of folio's
 * synthetic pad-width names, which is not an `ST_NumberFormat` member: the
 * custom spelling and its `w:format` string are restored by the splice the
 * save paths run (`selectiveXmlPatch.ts`), not reconstructed here, so the two
 * do not both claim to own it.
 */
const serializeNumFmt = (numFmt: NumberFormat): string =>
  `<w:numFmt w:val="${escapeXmlAttribute(numFmt)}"/>`;

/**
 * Serialize a level's paragraph properties — the modeled subset is indentation
 * plus tab stops (see `parseLevelParagraphProps`).
 *
 * An empty record is an empty `<w:pPr/>`, not an absent one: every child of
 * `CT_PPrGeneral` is optional, so the source may have written the element with
 * nothing in it and the record is what carries that.
 */
function serializeLevelParagraphProps(pPr: ParagraphFormatting): string {
  const indAttrs: string[] = [];
  if (pPr.indentLeft !== undefined) {
    indAttrs.push(`w:left="${intAttr(pPr.indentLeft)}"`);
  }
  if (pPr.indentRight !== undefined) {
    indAttrs.push(`w:right="${intAttr(pPr.indentRight)}"`);
  }
  if (pPr.indentFirstLine !== undefined) {
    if (pPr.hangingIndent) {
      indAttrs.push(`w:hanging="${intAttr(Math.abs(pPr.indentFirstLine))}"`);
    } else if (pPr.indentFirstLine !== 0) {
      indAttrs.push(`w:firstLine="${intAttr(pPr.indentFirstLine)}"`);
    }
  }

  const parts: string[] = [];
  if (pPr.tabs && pPr.tabs.length > 0) {
    const tabs = pPr.tabs
      .map((tab) => {
        const attrs = [`w:val="${tab.alignment}"`, `w:pos="${intAttr(tab.position)}"`];
        if (tab.leader) {
          attrs.push(`w:leader="${tab.leader}"`);
        }
        return `<w:tab ${attrs.join(" ")}/>`;
      })
      .join("");
    parts.push(`<w:tabs>${tabs}</w:tabs>`);
  }
  if (indAttrs.length > 0) {
    parts.push(`<w:ind ${indAttrs.join(" ")}/>`);
  }

  return parts.length === 0 ? "<w:pPr/>" : `<w:pPr>${parts.join("")}</w:pPr>`;
}

/**
 * Serialize one `w:lvl`. Children follow the ECMA-376 §17.9.6 CT_Lvl order
 * (start, numFmt, lvlRestart, pStyle, isLgl, suff, lvlText, lvlPicBulletId,
 * legacy, lvlJc, pPr, rPr) so a re-emitted level parses back into an
 * equivalent model.
 */
function serializeLevel(level: ListLevel): string {
  const parts: string[] = [];

  if (level.start !== undefined) {
    parts.push(`<w:start w:val="${intAttr(level.start)}"/>`);
  }
  parts.push(serializeNumFmt(level.numFmt));
  if (level.lvlRestart !== undefined) {
    parts.push(`<w:lvlRestart w:val="${intAttr(level.lvlRestart)}"/>`);
  }
  if (level.pStyle !== undefined) {
    parts.push(`<w:pStyle w:val="${escapeXmlAttribute(level.pStyle)}"/>`);
  }
  // An explicit `w:val="0"` turns legal numbering off where a container turned
  // it on, which a bare `<w:isLgl/>` would turn back on.
  if (level.isLgl !== undefined) {
    parts.push(level.isLgl ? "<w:isLgl/>" : '<w:isLgl w:val="0"/>');
  }
  if (level.suffix) {
    parts.push(`<w:suff w:val="${level.suffix}"/>`);
  }
  const lvlTextAttrs = [`w:val="${escapeXmlAttribute(level.lvlText)}"`];
  if (level.lvlTextNull !== undefined) {
    lvlTextAttrs.push(`w:null="${level.lvlTextNull ? 1 : 0}"`);
  }
  parts.push(`<w:lvlText ${lvlTextAttrs.join(" ")}/>`);
  if (level.lvlPicBulletId !== undefined) {
    parts.push(`<w:lvlPicBulletId w:val="${intAttr(level.lvlPicBulletId)}"/>`);
  }
  if (level.legacy) {
    const legacyAttrs: string[] = [];
    if (level.legacy.legacy !== undefined) {
      legacyAttrs.push(`w:legacy="${level.legacy.legacy ? 1 : 0}"`);
    }
    if (level.legacy.legacySpace !== undefined) {
      legacyAttrs.push(`w:legacySpace="${intAttr(level.legacy.legacySpace)}"`);
    }
    if (level.legacy.legacyIndent !== undefined) {
      legacyAttrs.push(`w:legacyIndent="${intAttr(level.legacy.legacyIndent)}"`);
    }
    parts.push(legacyAttrs.length === 0 ? "<w:legacy/>" : `<w:legacy ${legacyAttrs.join(" ")}/>`);
  }
  if (level.lvlJc) {
    parts.push(`<w:lvlJc w:val="${level.lvlJc}"/>`);
  }
  if (level.pPr) {
    parts.push(serializeLevelParagraphProps(level.pPr));
  }
  // A level's run properties reuse the run rPr serializer, so bullet fonts,
  // colors, and the vanish marker come out identical to body runs. An empty
  // record is an empty `<w:rPr/>`, which the source wrote and folio keeps.
  if (level.rPr) {
    parts.push(serializeTextFormatting(level.rPr) || "<w:rPr/>");
  }

  const attributes = [`w:ilvl="${intAttr(level.ilvl)}"`];
  if (level.tplc !== undefined) {
    attributes.push(`w:tplc="${escapeXmlAttribute(level.tplc)}"`);
  }
  if (level.tentative !== undefined) {
    attributes.push(`w:tentative="${level.tentative ? 1 : 0}"`);
  }
  const startTag = serializePreservedAttributes(attributes, level.preservedAttributes).join(" ");
  return `<w:lvl ${startTag}>${serializeWithPreservedChildren(parts, level.preserved)}</w:lvl>`;
}

/**
 * Serialize one `w:abstractNum` (the reusable list template). Children follow
 * ECMA-376 §17.9.1 CT_AbstractNum order (nsid, multiLevelType, tmpl, name,
 * styleLink, numStyleLink, lvl+).
 */
function serializeAbstractNum(abstractNum: AbstractNumbering): string {
  const parts: string[] = [];
  if (abstractNum.nsid !== undefined) {
    parts.push(`<w:nsid w:val="${escapeXmlAttribute(abstractNum.nsid)}"/>`);
  }
  if (abstractNum.multiLevelType) {
    parts.push(`<w:multiLevelType w:val="${abstractNum.multiLevelType}"/>`);
  }
  if (abstractNum.tmpl !== undefined) {
    parts.push(`<w:tmpl w:val="${escapeXmlAttribute(abstractNum.tmpl)}"/>`);
  }
  if (abstractNum.name !== undefined) {
    parts.push(`<w:name w:val="${escapeXmlAttribute(abstractNum.name)}"/>`);
  }
  if (abstractNum.styleLink !== undefined) {
    parts.push(`<w:styleLink w:val="${escapeXmlAttribute(abstractNum.styleLink)}"/>`);
  }
  if (abstractNum.numStyleLink !== undefined) {
    parts.push(`<w:numStyleLink w:val="${escapeXmlAttribute(abstractNum.numStyleLink)}"/>`);
  }
  const levels = [...abstractNum.levels].sort((a, b) => a.ilvl - b.ilvl);
  for (const level of levels) {
    parts.push(serializeLevel(level));
  }
  const attributes = serializePreservedAttributes(
    [`w:abstractNumId="${intAttr(abstractNum.abstractNumId)}"`],
    abstractNum.preservedAttributes,
  ).join(" ");
  const body = serializeWithPreservedChildren(parts, abstractNum.preserved);
  return `<w:abstractNum ${attributes}>${body}</w:abstractNum>`;
}

/** Serialize one `w:lvlOverride` (what an instance changes about one level). */
function serializeLevelOverride(override: LevelOverride): string {
  const parts: string[] = [];
  if (override.startOverride !== undefined) {
    parts.push(`<w:startOverride w:val="${intAttr(override.startOverride)}"/>`);
  }
  if (override.lvl) {
    parts.push(serializeLevel(override.lvl));
  }
  const attributes = serializePreservedAttributes(
    [`w:ilvl="${intAttr(override.ilvl)}"`],
    override.preservedAttributes,
  ).join(" ");
  const body = serializeWithPreservedChildren(parts, override.preserved);
  return `<w:lvlOverride ${attributes}>${body}</w:lvlOverride>`;
}

/**
 * Serialize one `w:num` (a concrete numbering instance referenced by `numId`).
 */
function serializeNum(instance: NumberingInstance): string {
  const parts: string[] = [`<w:abstractNumId w:val="${intAttr(instance.abstractNumId)}"/>`];
  for (const override of instance.levelOverrides ?? []) {
    parts.push(serializeLevelOverride(override));
  }
  const attributes = serializePreservedAttributes(
    [`w:numId="${intAttr(instance.numId)}"`],
    instance.preservedAttributes,
  ).join(" ");
  const body = serializeWithPreservedChildren(parts, instance.preserved);
  return `<w:num ${attributes}>${body}</w:num>`;
}

/**
 * Serialize {@link NumberingDefinitions} to a complete `word/numbering.xml`
 * string. `w:abstractNum` elements precede `w:num` elements as ECMA-376 §17.9
 * requires, and the part's own unmodelled children — a `w:numPicBullet`, the
 * `w:numIdMacAtCleanup` high-water mark — go back around them from the sink.
 */
export function serializeNumberingXml(numbering: NumberingDefinitions): string {
  const definitions = [
    ...numbering.abstractNums.map(serializeAbstractNum),
    ...numbering.nums.map(serializeNum),
  ];
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    serializePartElement({
      partPath: "word/numbering.xml",
      rootName: "w:numbering",
      rootAttributes: serializePreservedAttributes([], numbering.preservedAttributes).join(" "),
      baselinePrefixes: ["w"],
      sourceBindings: undefined,
      body: serializeWithPreservedChildren(definitions, numbering.preserved),
    })
  );
}
