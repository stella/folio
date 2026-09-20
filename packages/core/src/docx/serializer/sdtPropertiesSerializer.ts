/**
 * The property elements of a `w:sdt`, for every level that writes one.
 *
 * A block, inline, row-level and cell-level content control all open with
 * `w:sdtPr` and then `w:sdtEndPr`, and the rule for each was written twice
 * before a third level needed it. `w:sdtPr` is replayed when the parse
 * captured one and rebuilt from the modelled fields otherwise; only the
 * rebuild differs by level, so it is the caller's to supply.
 *
 * `w:sdtEndPr` is the part that is not a replay. `CT_SdtEndPr` declares
 * `w:rPr` and nothing else, so the model holds the element as a record and
 * its run properties as the one field, and a control folio rebuilds — one an
 * edit touched, or one a full repack writes — still carries its end mark.
 * Before that the element existed only as captured bytes and any rebuild
 * dropped it.
 */

import { escapeXmlAttribute } from "@stll/docx-core";

import type { SdtProperties } from "../../types/document";
import { reconcileRawSdtPr } from "../sdtPropertiesPatch";
import { serializeTextFormatting } from "./textFormattingSerializer";
import { isSingleWellFormedElement } from "./xmlUtils";

/**
 * A `w:sdtPr` built from the modelled fields, for a control with no captured
 * snapshot to replay: one a host constructed, or one an edit rebuilt.
 */
export function serializeFallbackSdtPr(props: SdtProperties): string {
  const parts: string[] = [];
  if (props.id !== undefined) {
    parts.push(`<w:id w:val="${props.id}"/>`);
  }
  if (props.alias) {
    parts.push(`<w:alias w:val="${escapeXmlAttribute(props.alias)}"/>`);
  }
  if (props.tag) {
    parts.push(`<w:tag w:val="${escapeXmlAttribute(props.tag)}"/>`);
  }
  if (props.lock) {
    parts.push(`<w:lock w:val="${props.lock}"/>`);
  }
  if (props.placeholder) {
    parts.push(
      `<w:placeholder><w:docPart w:val="${escapeXmlAttribute(props.placeholder)}"/></w:placeholder>`,
    );
  }
  if (props.showingPlaceholder) {
    parts.push("<w:showingPlcHdr/>");
  }
  // Type-specific child elements. Without these, a programmatically-
  // constructed control with `sdtType: "dropdown"` and a `listItems` set
  // would serialize as a bare `<w:sdtPr>` — Word would reopen the SDT as
  // richText and discard the dropdown items. `reconcileRawSdtPr` only
  // patches existing markers (it does not insert a missing
  // `<w:dropDownList>`), so the fallback must emit the type-defining
  // marker itself.
  switch (props.sdtType) {
    case "plainText":
      parts.push("<w:text/>");
      break;
    case "date": {
      const fullDateAttr = props.dateValueISO
        ? ` w:fullDate="${escapeXmlAttribute(props.dateValueISO)}"`
        : "";
      const formatChild = props.dateFormat
        ? `<w:dateFormat w:val="${escapeXmlAttribute(props.dateFormat)}"/>`
        : "";
      if (fullDateAttr || formatChild) {
        parts.push(`<w:date${fullDateAttr}>${formatChild}</w:date>`);
      } else {
        parts.push("<w:date/>");
      }
      break;
    }
    case "dropdown":
    case "comboBox": {
      const tag = props.sdtType === "dropdown" ? "w:dropDownList" : "w:comboBox";
      const items = (props.listItems ?? [])
        .map(
          (item) =>
            `<w:listItem w:displayText="${escapeXmlAttribute(item.displayText)}" w:value="${escapeXmlAttribute(item.value)}"/>`,
        )
        .join("");
      parts.push(`<${tag}>${items}</${tag}>`);
      break;
    }
    case "checkbox": {
      const val = props.checked ? "1" : "0";
      parts.push(`<w14:checkbox><w14:checked w14:val="${val}"/></w14:checkbox>`);
      break;
    }
    case "picture":
      parts.push("<w:picture/>");
      break;
    case "buildingBlockGallery":
      parts.push("<w:docPartObj/>");
      break;
    case "group":
      parts.push("<w:group/>");
      break;
    default:
      // richText / unknown — no specific marker; bare <w:sdtPr> means
      // richText per the OOXML default.
      break;
  }
  return `<w:sdtPr>${parts.join("")}</w:sdtPr>`;
}

/**
 * `w:sdtEndPr`, replayed when the parse captured it and written from the
 * record otherwise.
 *
 * The record's presence is the element's: `<w:sdtEndPr/>` with no `w:rPr` is
 * what Word writes for most controls and says something a missing element
 * does not.
 */
export const serializeSdtEndProperties = (properties: SdtProperties): string => {
  if (
    properties.rawEndPropertiesXml &&
    isSingleWellFormedElement(properties.rawEndPropertiesXml, "sdtEndPr")
  ) {
    return properties.rawEndPropertiesXml;
  }
  const endProperties = properties.endProperties;
  if (endProperties === undefined) {
    return "";
  }
  const runProperties = serializeTextFormatting(endProperties.runProperties);
  return runProperties.length > 0 ? `<w:sdtEndPr>${runProperties}</w:sdtEndPr>` : "<w:sdtEndPr/>";
};

/**
 * `w:sdtPr` followed by `w:sdtEndPr`, in the order `CT_Sdt*` declares them.
 *
 * Replay the captured `w:sdtPr` only when it is structurally a single
 * `<w:sdtPr>` element — a malformed or attacker-supplied string (one that
 * closes `<w:sdt>` early, say, or injects sibling markup) falls back to the
 * rebuilt properties instead of being spliced into the document verbatim.
 * Whatever the source, the modelled interactive edits (a checkbox toggle, a
 * date pick, a dropdown selection) are reconciled into it, so an edit is not
 * discarded by a replay and unmodelled markers inside the raw string are left
 * untouched.
 *
 * `properties.dropdownLastValue` is the only record of what was selected: the
 * XSD default of `@w:lastValue` is the empty string, so "never selected",
 * "cleared" and "selected" are three distinguishable states, and the body's
 * display text is evidence for none of them — it is equally the placeholder of
 * a dropdown nobody has touched, and a `displayText` shared by two list items
 * picks the wrong sibling. `""` is a value a producer can author
 * (`<w:listItem w:value=""/>`), so presence is the test, not truthiness.
 * `properties.dateValueISO` is the same for a date: the body shows the
 * format-rendered display ("2 June 2026" per `dateFormat`), which is not the
 * ISO 8601 `w:fullDate` requires.
 *
 * @param fallbackPropertiesXml the `w:sdtPr` to write when no captured one can
 *   be replayed; the inline level spells a different set of fields, so it
 *   passes its own.
 */
export const serializeSdtPropertyElements = (
  properties: SdtProperties,
  fallbackPropertiesXml: string = serializeFallbackSdtPr(properties),
): string => {
  const basePropertiesXml =
    properties.rawPropertiesXml && isSingleWellFormedElement(properties.rawPropertiesXml, "sdtPr")
      ? properties.rawPropertiesXml
      : fallbackPropertiesXml;
  const dateFullDate =
    properties.sdtType === "date" && properties.dateValueISO ? properties.dateValueISO : undefined;
  const dropdownLastValue =
    (properties.sdtType === "dropdown" || properties.sdtType === "comboBox") &&
    typeof properties.dropdownLastValue === "string"
      ? properties.dropdownLastValue
      : undefined;
  const sdtPrXml = reconcileRawSdtPr(basePropertiesXml, properties, {
    ...(dateFullDate !== undefined ? { dateFullDate } : {}),
    ...(dropdownLastValue !== undefined ? { dropdownLastValue } : {}),
  });
  return `${sdtPrXml}${serializeSdtEndProperties(properties)}`;
};
