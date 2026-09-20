/**
 * The one writer of `<w:sdtPr>`.
 *
 * Every content control — block, inline, row and cell — writes its property
 * set here. There used to be three producers: a fallback projection on the
 * block path, a second one on the inline path, and a regex pass over the
 * source's captured bytes that patched an interactive change back into them.
 * The three agreed on nothing in particular, and the patching pass existed
 * only because the replay was the serialization source.
 *
 * It is not any more. `parseSdtProperties` puts every child folio does not
 * model into `SdtProperties.preserved` at its `CT_SdtPr` ordinal, so the model
 * holds the whole element and the writer merges its two halves back in schema
 * order — which `CT_SdtPr` demands, and Word repairs a file that disregards.
 */

import { escapeXmlAttribute, serializeOnOffElement } from "@stll/docx-core";
import { assertSafePreservedMarkup, serializeSequenceChildren } from "@stll/docx-core/schema";

import type { SdtProperties } from "../../types/document";
import type { DeclaredChild } from "../containerChildren.gen";
import { statesControlKind } from "../sdtProperties";
import { withModelledControlState } from "../sdtPropertiesPatch";
import { serializeTextFormatting } from "./textFormattingSerializer";
import { isSingleWellFormedElement } from "./xmlUtils";

/**
 * One child of the set, and the sequence slot it belongs in.
 *
 * The slot is named by a declared child rather than given as a number, so the
 * order comes from the generated sequence instead of from a second table here.
 */
type PlacedChild = readonly [slot: DeclaredChild<"content-control-properties">, xml: string];

const listElement = (name: "dropDownList" | "comboBox", props: SdtProperties): string => {
  const lastValue =
    props.dropdownLastValue === undefined
      ? ""
      : ` w:lastValue="${escapeXmlAttribute(props.dropdownLastValue)}"`;
  const items = (props.listItems ?? [])
    .map(
      (item) =>
        `<w:listItem w:displayText="${escapeXmlAttribute(item.displayText)}" w:value="${escapeXmlAttribute(item.value)}"/>`,
    )
    .join("");
  return `<w:${name}${lastValue}>${items}</w:${name}>`;
};

/**
 * The kind marker for a control that carries none of its own.
 *
 * Only reached for a control built in code rather than parsed: a parsed one
 * keeps the element its author wrote, in the sink, glyph elements and all.
 * `richText` is the format's default and `unknown` is folio's word for a kind
 * it could not read, so neither writes a marker — a bare `w:sdtPr` already
 * means richText.
 *
 * A checkbox has no declared kind element at all: `w14:checkbox` is an
 * extension the Transitional content model does not know, so it is filed
 * under the slot the kind choice occupies, which is where Word writes it.
 */
const synthesizedKind = (props: SdtProperties): PlacedChild => {
  switch (props.sdtType) {
    case "plainText":
      return ["text", "<w:text/>"];
    case "date": {
      const fullDate = props.dateValueISO
        ? ` w:fullDate="${escapeXmlAttribute(props.dateValueISO)}"`
        : "";
      const format = props.dateFormat
        ? `<w:dateFormat w:val="${escapeXmlAttribute(props.dateFormat)}"/>`
        : "";
      return ["date", `<w:date${fullDate}>${format}</w:date>`];
    }
    case "dropdown":
      return ["dropDownList", listElement("dropDownList", props)];
    case "comboBox":
      return ["comboBox", listElement("comboBox", props)];
    case "checkbox":
      return [
        "equation",
        `<w14:checkbox><w14:checked w14:val="${props.checked ? "1" : "0"}"/></w14:checkbox>`,
      ];
    case "picture":
      return ["picture", "<w:picture/>"];
    case "buildingBlockGallery":
      return ["docPartObj", "<w:docPartObj/>"];
    case "group":
      return ["group", "<w:group/>"];
    case "richText":
    case "unknown":
      return ["richText", ""];
    default: {
      const exhaustive: never = props.sdtType;
      return exhaustive;
    }
  }
};

/**
 * `<w:sdtPr>` from the model: the modelled children, the preserved ones, in
 * schema order.
 *
 * An empty property set is written as an empty element rather than left out.
 * `<w:sdtPr/>` and no `w:sdtPr` at all are different documents — the first is
 * a richText control that states nothing, the second is markup the schema does
 * not admit under `w:sdt` — and folio used to write the second for both.
 */
export const serializeSdtProperties = (props: SdtProperties): string => {
  assertSafePreservedMarkup(props.preserved);
  const modelled: PlacedChild[] = [];
  if (props.alias !== undefined) {
    modelled.push(["alias", `<w:alias w:val="${escapeXmlAttribute(props.alias)}"/>`]);
  }
  if (props.tag !== undefined) {
    modelled.push(["tag", `<w:tag w:val="${escapeXmlAttribute(props.tag)}"/>`]);
  }
  if (props.id !== undefined) {
    modelled.push(["id", `<w:id w:val="${props.id}"/>`]);
  }
  if (props.lock !== undefined) {
    modelled.push(["lock", `<w:lock w:val="${props.lock}"/>`]);
  }
  if (props.placeholder !== undefined) {
    modelled.push([
      "placeholder",
      `<w:placeholder><w:docPart w:val="${escapeXmlAttribute(props.placeholder)}"/></w:placeholder>`,
    ]);
  }
  if (props.showingPlaceholder !== undefined) {
    // Tri-state: absent, on, off. An explicit `w:val="0"` is what the source
    // wrote, and it is not the same statement as writing nothing.
    modelled.push([
      "showingPlcHdr",
      serializeOnOffElement(props.showingPlaceholder, "showingPlcHdr"),
    ]);
  }
  // The kind element is kept as bytes, but four of its values are modelled
  // because a user changes them; the model wins for those and the element
  // keeps everything else. See `sdtPropertiesPatch.ts`.
  const captured = props.preserved?.children ?? [];
  let statedKind = false;
  const preserved = captured.map((child) => {
    if (!statesControlKind(child.xml)) {
      return child;
    }
    statedKind = true;
    return { index: child.index, xml: withModelledControlState(child.xml, props) };
  });
  if (!statedKind) {
    modelled.push(synthesizedKind(props));
  }

  const children = serializeSequenceChildren({
    container: "content-control-properties",
    modelled,
    preserved: props.preserved === undefined ? undefined : { children: preserved },
  });
  const body = children.join("");
  return body.length === 0 ? "<w:sdtPr/>" : `<w:sdtPr>${body}</w:sdtPr>`;
};

/**
 * `w:sdtEndPr`, replayed when the parse captured it and written from the
 * record otherwise.
 *
 * The record's presence is the element's: `<w:sdtEndPr/>` with no `w:rPr` is
 * what Word writes for most controls and differs from a missing element.
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

/** `w:sdtPr` followed by `w:sdtEndPr`, in schema order. */
export const serializeSdtPropertyElements = (properties: SdtProperties): string =>
  `${serializeSdtProperties(properties)}${serializeSdtEndProperties(properties)}`;

/** The one writer of an SDT wrapper at every content-model level. */
export const serializeSdtWrapper = (properties: SdtProperties, contentXml: string): string => {
  const propertyElements = serializeSdtPropertyElements(properties);
  const before = properties.rawSdtChildrenBeforeContent ?? "";
  const after = properties.rawSdtChildrenAfterContent ?? "";
  return `<w:sdt>${propertyElements}${before}<w:sdtContent>${contentXml}</w:sdtContent>${after}</w:sdt>`;
};
