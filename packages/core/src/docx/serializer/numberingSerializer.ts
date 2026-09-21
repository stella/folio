/**
 * Numbering Serializer - Serialize numbering definitions back to OOXML XML
 *
 * Converts the parsed {@link NumberingDefinitions} model (abstract numberings
 * and concrete numbering instances) back into `word/numbering.xml`.
 *
 * Every container writes its children in the order its content model declares
 * through the same generated sequence that its parser uses for capture
 * positions. Unmodelled markup therefore returns to its schema slot.
 *
 * Save paths still splice definitions rather than overwrite the whole part:
 * they compare this output with a re-serialization of the original model and
 * retain byte-exact source definitions that the editor did not change.
 */

import type {
  AbstractNumbering,
  LevelOverride,
  ListLevel,
  NumberingDefinitions,
  NumberingInstance,
} from "../../types/document";
import { serializeParagraphPropertySet } from "../../internal/paragraphFormattingSerialization";
import { serializePreservedAttributes } from "../attributeRemainder";
import { serializePartElement } from "./partNamespaces";
import { serializeTextFormatting } from "./textFormattingSerializer";
import { intAttr } from "./xmlUtils";
import { escapeXmlAttribute, serializeOnOffElement } from "@stll/docx-core";
import { serializeSequenceChildren } from "@stll/docx-core/schema";

/** Serialize one `w:lvl`. */
function serializeLevel(level: ListLevel): string {
  const numFmtFormat =
    level.numFmtFormat === undefined
      ? ""
      : ` w:format="${escapeXmlAttribute(level.numFmtFormat)}"`;
  const pPr =
    level.pPr === undefined
      ? ""
      : serializeParagraphPropertySet({ formatting: level.pPr }) || "<w:pPr/>";
  const rPr = level.rPr === undefined ? "" : serializeTextFormatting(level.rPr) || "<w:rPr/>";
  const lvlTextAttrs = [`w:val="${escapeXmlAttribute(level.lvlText)}"`];
  if (level.lvlTextNull !== undefined) {
    lvlTextAttrs.push(`w:null="${level.lvlTextNull ? 1 : 0}"`);
  }

  let legacy = "";
  if (level.legacy !== undefined) {
    const attributes: string[] = [];
    if (level.legacy.legacy !== undefined) {
      attributes.push(`w:legacy="${level.legacy.legacy ? 1 : 0}"`);
    }
    if (level.legacy.legacySpace !== undefined) {
      attributes.push(`w:legacySpace="${intAttr(level.legacy.legacySpace)}"`);
    }
    if (level.legacy.legacyIndent !== undefined) {
      attributes.push(`w:legacyIndent="${intAttr(level.legacy.legacyIndent)}"`);
    }
    legacy = attributes.length === 0 ? "<w:legacy/>" : `<w:legacy ${attributes.join(" ")}/>`;
  }

  const children = serializeSequenceChildren({
    container: "w:lvl",
    modelled: [
      ["start", level.start === undefined ? "" : `<w:start w:val="${intAttr(level.start)}"/>`],
      ["numFmt", `<w:numFmt w:val="${escapeXmlAttribute(level.numFmt)}"${numFmtFormat}/>`],
      [
        "lvlRestart",
        level.lvlRestart === undefined
          ? ""
          : `<w:lvlRestart w:val="${intAttr(level.lvlRestart)}"/>`,
      ],
      [
        "pStyle",
        level.pStyle === undefined ? "" : `<w:pStyle w:val="${escapeXmlAttribute(level.pStyle)}"/>`,
      ],
      ["isLgl", serializeOnOffElement(level.isLgl, "isLgl")],
      [
        "suff",
        level.suffix === undefined ? "" : `<w:suff w:val="${escapeXmlAttribute(level.suffix)}"/>`,
      ],
      ["lvlText", `<w:lvlText ${lvlTextAttrs.join(" ")}/>`],
      [
        "lvlPicBulletId",
        level.lvlPicBulletId === undefined
          ? ""
          : `<w:lvlPicBulletId w:val="${intAttr(level.lvlPicBulletId)}"/>`,
      ],
      ["legacy", legacy],
      [
        "lvlJc",
        level.lvlJc === undefined ? "" : `<w:lvlJc w:val="${escapeXmlAttribute(level.lvlJc)}"/>`,
      ],
      ["pPr", pPr],
      ["rPr", rPr],
    ],
    preserved: level.preserved,
  });

  const attributes = [`w:ilvl="${intAttr(level.ilvl)}"`];
  if (level.tplc !== undefined) {
    attributes.push(`w:tplc="${escapeXmlAttribute(level.tplc)}"`);
  }
  if (level.tentative !== undefined) {
    attributes.push(`w:tentative="${level.tentative ? 1 : 0}"`);
  }
  const startTag = serializePreservedAttributes(attributes, level.preservedAttributes).join(" ");
  return `<w:lvl ${startTag}>${children.join("")}</w:lvl>`;
}

/** Serialize one reusable `w:abstractNum` list template. */
function serializeAbstractNum(abstractNum: AbstractNumbering): string {
  const levels = [...abstractNum.levels].sort((left, right) => left.ilvl - right.ilvl);
  const children = serializeSequenceChildren({
    container: "w:abstractNum",
    modelled: [
      [
        "nsid",
        abstractNum.nsid === undefined
          ? ""
          : `<w:nsid w:val="${escapeXmlAttribute(abstractNum.nsid)}"/>`,
      ],
      [
        "multiLevelType",
        abstractNum.multiLevelType === undefined
          ? ""
          : `<w:multiLevelType w:val="${escapeXmlAttribute(abstractNum.multiLevelType)}"/>`,
      ],
      [
        "tmpl",
        abstractNum.tmpl === undefined
          ? ""
          : `<w:tmpl w:val="${escapeXmlAttribute(abstractNum.tmpl)}"/>`,
      ],
      [
        "name",
        abstractNum.name === undefined
          ? ""
          : `<w:name w:val="${escapeXmlAttribute(abstractNum.name)}"/>`,
      ],
      [
        "styleLink",
        abstractNum.styleLink === undefined
          ? ""
          : `<w:styleLink w:val="${escapeXmlAttribute(abstractNum.styleLink)}"/>`,
      ],
      [
        "numStyleLink",
        abstractNum.numStyleLink === undefined
          ? ""
          : `<w:numStyleLink w:val="${escapeXmlAttribute(abstractNum.numStyleLink)}"/>`,
      ],
      ["lvl", levels.map(serializeLevel).join("")],
    ],
    preserved: abstractNum.preserved,
  });
  const attributes = serializePreservedAttributes(
    [`w:abstractNumId="${intAttr(abstractNum.abstractNumId)}"`],
    abstractNum.preservedAttributes,
  ).join(" ");
  return `<w:abstractNum ${attributes}>${children.join("")}</w:abstractNum>`;
}

/** Serialize one concrete instance's per-level override. */
function serializeLevelOverride(override: LevelOverride): string {
  const children = serializeSequenceChildren({
    container: "w:lvlOverride",
    modelled: [
      [
        "startOverride",
        override.startOverride === undefined
          ? ""
          : `<w:startOverride w:val="${intAttr(override.startOverride)}"/>`,
      ],
      ["lvl", override.lvl === undefined ? "" : serializeLevel(override.lvl)],
    ],
    preserved: override.preserved,
  });
  const attributes = serializePreservedAttributes(
    [`w:ilvl="${intAttr(override.ilvl)}"`],
    override.preservedAttributes,
  ).join(" ");
  return `<w:lvlOverride ${attributes}>${children.join("")}</w:lvlOverride>`;
}

/** Serialize one concrete `w:num` numbering instance. */
function serializeNum(instance: NumberingInstance): string {
  const children = serializeSequenceChildren({
    container: "w:num",
    modelled: [
      ["abstractNumId", `<w:abstractNumId w:val="${intAttr(instance.abstractNumId)}"/>`],
      ["lvlOverride", (instance.levelOverrides ?? []).map(serializeLevelOverride).join("")],
    ],
    preserved: instance.preserved,
  });
  const attributes = serializePreservedAttributes(
    [`w:numId="${intAttr(instance.numId)}"`],
    instance.preservedAttributes,
  ).join(" ");
  return `<w:num ${attributes}>${children.join("")}</w:num>`;
}

/** Serialize a complete `word/numbering.xml` part. */
export function serializeNumberingXml(numbering: NumberingDefinitions): string {
  const children = serializeSequenceChildren({
    container: "w:numbering",
    modelled: [
      ["abstractNum", numbering.abstractNums.map(serializeAbstractNum).join("")],
      ["num", numbering.nums.map(serializeNum).join("")],
    ],
    preserved: numbering.preserved,
  });
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    serializePartElement({
      partPath: "word/numbering.xml",
      rootName: "w:numbering",
      rootAttributes: serializePreservedAttributes([], numbering.preservedAttributes).join(" "),
      baselinePrefixes: ["w"],
      sourceBindings: undefined,
      body: children.join(""),
    })
  );
}
