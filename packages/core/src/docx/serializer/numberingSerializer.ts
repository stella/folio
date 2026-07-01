/**
 * Numbering Serializer - Serialize numbering definitions back to OOXML XML
 *
 * Converts the parsed {@link NumberingDefinitions} model (abstract numberings
 * and concrete numbering instances) back into `word/numbering.xml` fragments.
 * The inverse of `numberingParser`.
 *
 * OOXML Reference:
 * - Root: w:numbering
 * - Templates: w:abstractNum[@w:abstractNumId] with 0..8 w:lvl children
 * - Instances: w:num[@w:numId] referencing an abstractNum (+ optional overrides)
 *
 * The document model does NOT faithfully retain everything `numbering.xml`
 * carries: abstract numberings drop `w:nsid` / `w:tmpl`, custom number formats
 * (`w:numFmt w:val="custom"` wrapped in mc:AlternateContent) collapse to a
 * decimalZero family value, and a level's `w:pPr` / `w:rPr` keep only the subset
 * the parser models. Callers therefore must NOT overwrite the whole part with
 * this output — it would drop those pieces. The save paths use this serializer
 * only to detect which `w:abstractNum` / `w:num` definitions the model actually
 * changed (by comparing against a re-parse+re-serialize baseline so the lossy
 * parse cancels out) and to splice just those definitions into the original
 * part, keeping every untouched definition and sub-element byte-exact.
 */

import type {
  AbstractNumbering,
  ListLevel,
  NumberingDefinitions,
  NumberingInstance,
  ParagraphFormatting,
} from "../../types/document";
import { serializeTextFormatting } from "./runSerializer";
import { escapeXml, intAttr } from "./xmlUtils";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/**
 * Serialize a level's paragraph properties — the modeled subset is indentation
 * plus tab stops (see `parseLevelParagraphProps`). Returns "" when neither is
 * present so an empty `<w:pPr/>` is not emitted.
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

  if (parts.length === 0) {
    return "";
  }
  return `<w:pPr>${parts.join("")}</w:pPr>`;
}

/**
 * Serialize one `w:lvl`. Children follow the ECMA-376 §17.9.6 CT_Lvl order
 * (start, numFmt, lvlRestart, isLgl, suff, lvlText, legacy, lvlJc, pPr, rPr) so
 * a re-emitted level parses back into an equivalent model.
 */
function serializeLevel(level: ListLevel): string {
  const parts: string[] = [];

  if (level.start !== undefined) {
    parts.push(`<w:start w:val="${intAttr(level.start)}"/>`);
  }
  parts.push(`<w:numFmt w:val="${escapeXml(level.numFmt)}"/>`);
  if (level.lvlRestart !== undefined) {
    parts.push(`<w:lvlRestart w:val="${intAttr(level.lvlRestart)}"/>`);
  }
  if (level.isLgl) {
    parts.push("<w:isLgl/>");
  }
  if (level.suffix) {
    parts.push(`<w:suff w:val="${level.suffix}"/>`);
  }
  parts.push(`<w:lvlText w:val="${escapeXml(level.lvlText)}"/>`);
  if (level.legacy) {
    const legacyAttrs: string[] = [`w:legacy="${level.legacy.legacy ? 1 : 0}"`];
    if (level.legacy.legacySpace !== undefined) {
      legacyAttrs.push(`w:legacySpace="${intAttr(level.legacy.legacySpace)}"`);
    }
    if (level.legacy.legacyIndent !== undefined) {
      legacyAttrs.push(`w:legacyIndent="${intAttr(level.legacy.legacyIndent)}"`);
    }
    parts.push(`<w:legacy ${legacyAttrs.join(" ")}/>`);
  }
  if (level.lvlJc) {
    parts.push(`<w:lvlJc w:val="${level.lvlJc}"/>`);
  }
  if (level.pPr) {
    parts.push(serializeLevelParagraphProps(level.pPr));
  }
  // A level's run properties reuse the run rPr serializer, so bullet fonts,
  // colors, and the vanish marker come out identical to body runs.
  parts.push(serializeTextFormatting(level.rPr));

  return `<w:lvl w:ilvl="${intAttr(level.ilvl)}">${parts.join("")}</w:lvl>`;
}

/**
 * Serialize one `w:abstractNum` (the reusable list template). Children follow
 * ECMA-376 §17.9.1 CT_AbstractNum order for the modeled subset (multiLevelType,
 * name, styleLink, numStyleLink, lvl+); `w:nsid` / `w:tmpl` are not modeled.
 */
function serializeAbstractNum(abstractNum: AbstractNumbering): string {
  const parts: string[] = [];
  if (abstractNum.multiLevelType) {
    parts.push(`<w:multiLevelType w:val="${abstractNum.multiLevelType}"/>`);
  }
  if (abstractNum.name !== undefined) {
    parts.push(`<w:name w:val="${escapeXml(abstractNum.name)}"/>`);
  }
  if (abstractNum.styleLink !== undefined) {
    parts.push(`<w:styleLink w:val="${escapeXml(abstractNum.styleLink)}"/>`);
  }
  if (abstractNum.numStyleLink !== undefined) {
    parts.push(`<w:numStyleLink w:val="${escapeXml(abstractNum.numStyleLink)}"/>`);
  }
  const levels = [...abstractNum.levels].sort((a, b) => a.ilvl - b.ilvl);
  for (const level of levels) {
    parts.push(serializeLevel(level));
  }
  return `<w:abstractNum w:abstractNumId="${intAttr(abstractNum.abstractNumId)}">${parts.join("")}</w:abstractNum>`;
}

/**
 * Serialize one `w:num` (a concrete numbering instance referenced by `numId`).
 */
function serializeNum(instance: NumberingInstance): string {
  const parts: string[] = [`<w:abstractNumId w:val="${intAttr(instance.abstractNumId)}"/>`];
  for (const override of instance.levelOverrides ?? []) {
    const overrideParts: string[] = [];
    if (override.startOverride !== undefined) {
      overrideParts.push(`<w:startOverride w:val="${intAttr(override.startOverride)}"/>`);
    }
    if (override.lvl) {
      overrideParts.push(serializeLevel(override.lvl));
    }
    parts.push(
      `<w:lvlOverride w:ilvl="${intAttr(override.ilvl)}">${overrideParts.join("")}</w:lvlOverride>`,
    );
  }
  return `<w:num w:numId="${intAttr(instance.numId)}">${parts.join("")}</w:num>`;
}

/**
 * Serialize {@link NumberingDefinitions} to a complete `word/numbering.xml`
 * string. `w:abstractNum` elements precede `w:num` elements as ECMA-376 §17.9
 * requires. See the module note: this whole-part output is used only for
 * change detection and per-definition splicing, never to overwrite the part.
 */
export function serializeNumberingXml(numbering: NumberingDefinitions): string {
  const abstractNums = numbering.abstractNums.map(serializeAbstractNum).join("");
  const nums = numbering.nums.map(serializeNum).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:numbering xmlns:w="${W_NS}">${abstractNums}${nums}</w:numbering>`;
}
