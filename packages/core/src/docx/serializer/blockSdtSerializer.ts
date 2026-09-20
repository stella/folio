/**
 * Block-level SDT serializer.
 *
 * Replays `<w:sdtPr>` (and `<w:sdtEndPr>` when present) verbatim from the
 * `rawPropertiesXml` / `rawEndPropertiesXml` strings captured by the parser.
 * That keeps OOXML element order intact (`CT_SdtPr` is an `xsd:sequence`,
 * ECMA-376 §17.5.2) and round-trips unmodeled features — `w:dataBinding`,
 * `w15:repeatingSection`, `@lastValue`, custom XML mappings — without us
 * having to enumerate them.
 *
 * If a `BlockSdt` was constructed programmatically (no parsed snapshot to
 * replay), fall back to a minimal projection from the modeled fields so the
 * result is still a valid `<w:sdt>`.
 *
 * Sharing the helper between the document body and the header/footer
 * serializers keeps body↔HF parity in one place.
 */

import { escapeXmlAttribute } from "@stll/docx-core";

import type { BlockContent, BlockSdt, SdtProperties } from "../../types/document";
import { reconcileRawSdtPr } from "../sdtPropertiesPatch";
import { isSingleWellFormedElement } from "./xmlUtils";

/**
 * Synthesize a `<w:sdtPr>` from the modeled {@link SdtProperties}.
 *
 * Reached for any SDT — block, row, cell or inline — carrying no captured
 * `rawPropertiesXml`, so it was constructed programmatically rather than
 * parsed. One builder, because the two spellings that used to differ have one
 * answer each:
 *
 *  - `w:id/@w:val` is `ST_DecimalNumber` (ECMA-376 §17.5.2.18, §22.9.2.3),
 *    i.e. `xsd:integer`. Anything else is not a spelling of the attribute, so
 *    write it only when the id is an integer.
 *  - `w:lock` is absent by default and Word writes it only when something is
 *    locked, so `unlocked` is the absence, not a value to emit. An authored
 *    `<w:lock w:val="unlocked"/>` is not lost by this: it lives in
 *    `rawPropertiesXml` and never reaches this builder.
 */
export function serializeFallbackSdtPr(props: SdtProperties): string {
  const parts: string[] = [];
  if (Number.isInteger(props.id)) {
    parts.push(`<w:id w:val="${String(props.id)}"/>`);
  }
  if (props.alias) {
    parts.push(`<w:alias w:val="${escapeXmlAttribute(props.alias)}"/>`);
  }
  if (props.tag) {
    parts.push(`<w:tag w:val="${escapeXmlAttribute(props.tag)}"/>`);
  }
  if (props.lock && props.lock !== "unlocked") {
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
 * `properties.dropdownLastValue` is the only record of what was selected.
 *
 * The XSD default of `@w:lastValue` is the empty string, so "never selected",
 * "cleared" and "selected" are three distinguishable states, and the body's
 * display text is evidence for none of them: it is equally the placeholder of
 * a dropdown nobody has touched, and a displayText shared by two list items
 * picks the wrong sibling. The parser records the authored `@w:lastValue` and
 * `setContentControlValue` records a pick; what neither wrote is no selection.
 *
 * `""` is a value a producer can author (`<w:listItem w:value=""/>`), so
 * presence is the test, not truthiness.
 */
function extractDropdownLastValue(blockSdt: BlockSdt): string | undefined {
  if (blockSdt.properties.sdtType !== "dropdown" && blockSdt.properties.sdtType !== "comboBox") {
    return undefined;
  }
  return blockSdt.properties.dropdownLastValue;
}

function extractDateFullDate(blockSdt: BlockSdt): string | undefined {
  if (blockSdt.properties.sdtType !== "date") {
    return undefined;
  }
  // The ISO bound value lives on the modeled `dateValueISO`. We deliberately
  // do NOT read the SDT body — the body shows the format-rendered display
  // ("2 June 2026" per dateFormat "d MMMM yyyy") which would corrupt
  // `w:fullDate` (which OOXML requires to be ISO 8601). If the model has no
  // ISO value yet (e.g. a fresh control the user never picked a date for),
  // omit `w:fullDate` so the serializer doesn't write a garbage one.
  const iso = blockSdt.properties.dateValueISO;
  return iso !== undefined && iso.length > 0 ? iso : undefined;
}

export function serializeBlockSdt(
  blockSdt: BlockSdt,
  serializeChild: (block: BlockContent) => string,
): string {
  const props = blockSdt.properties;
  // Replay the captured snapshot only when it is structurally a single
  // `<w:sdtPr>` element — a malformed or attacker-supplied string (e.g. one
  // that closes `<w:sdt>` early or injects sibling markup) falls back to a
  // synthesized properties block instead of being spliced into the document
  // verbatim.
  const baseSdtPr =
    props.rawPropertiesXml && isSingleWellFormedElement(props.rawPropertiesXml, "sdtPr")
      ? props.rawPropertiesXml
      : serializeFallbackSdtPr(props);
  // Reconcile any modeled property mutations the editor may have made into
  // the raw XML before replay so checkbox / dropdown / date interactions
  // survive the round-trip. Unmodeled markers (dataBinding,
  // w15:repeatingSection, etc.) inside the raw string are preserved.
  const dateFullDate = extractDateFullDate(blockSdt);
  const dropdownLastValue = extractDropdownLastValue(blockSdt);
  const sdtPrXml = reconcileRawSdtPr(baseSdtPr, props, {
    ...(dateFullDate !== undefined ? { dateFullDate } : {}),
    ...(dropdownLastValue !== undefined ? { dropdownLastValue } : {}),
  });
  const sdtEndPrXml =
    props.rawEndPropertiesXml && isSingleWellFormedElement(props.rawEndPropertiesXml, "sdtEndPr")
      ? props.rawEndPropertiesXml
      : "";
  const contentXml = blockSdt.content.map(serializeChild).join("");
  // Replay any direct sdt children that lived OUTSIDE sdtContent at parse
  // time (range markers per MS-OE376 §2.5.2.30: bookmark / comment /
  // tracked-change / custom XML ranges that span an SDT boundary). Position
  // matters — the captured before/after strings preserve which side of
  // sdtContent each marker sat on.
  const before = props.rawSdtChildrenBeforeContent ?? "";
  const after = props.rawSdtChildrenAfterContent ?? "";
  return `<w:sdt>${sdtPrXml}${sdtEndPrXml}${before}<w:sdtContent>${contentXml}</w:sdtContent>${after}</w:sdt>`;
}
