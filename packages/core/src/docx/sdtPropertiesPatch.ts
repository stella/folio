/**
 * The modelled interactive state, written into the control-kind element.
 *
 * `w:date`, `w:dropDownList`, `w:comboBox` and `w14:checkbox` are kept as
 * bytes: each carries children and attributes {@link SdtProperties} has no
 * field for — a calendar and a locale, a gallery, the two checkbox glyphs —
 * and rebuilding one from the model would drop them. But four of their values
 * *are* modelled, because a user changes them: the checkbox state, the bound
 * date, the display format and the chosen list value. The model wins for
 * those four; the element keeps everything else.
 *
 * This is the one place the two halves meet, and it runs inside the one
 * `w:sdtPr` writer rather than beside it.
 */

import { escapeXmlAttribute } from "@stll/docx-core";

import type { SdtProperties } from "../types/document";

/**
 * Drop any `*:lastValue="…"` attribute (any namespace prefix, or unprefixed)
 * from an attribute-list string.
 *
 * A scan rather than one greedy regex, which lint flags as backtracking-risky.
 * Prefix tolerance is the point: a source that binds the Word namespace under
 * `ns0` writes `ns0:lastValue`, and keeping it beside a freshly emitted one
 * would leave the element with two.
 */
function stripLastValueAttr(attrs: string): string {
  const SUFFIX = "lastValue=";
  let out = attrs;
  let searchFrom = 0;
  while (searchFrom < out.length) {
    const hitIdx = out.indexOf(SUFFIX, searchFrom);
    if (hitIdx === -1) {
      break;
    }
    // Walk backwards over an optional `prefix:` and require a leading
    // whitespace separator so an attribute like `mylastValue=` is left alone.
    let nameStart = hitIdx;
    while (nameStart > 0 && /[A-Za-z0-9_-]/u.test(out[nameStart - 1] ?? "")) {
      nameStart -= 1;
    }
    if (nameStart > 0 && out[nameStart - 1] === ":") {
      nameStart -= 1;
      while (nameStart > 0 && /[A-Za-z0-9_-]/u.test(out[nameStart - 1] ?? "")) {
        nameStart -= 1;
      }
    }
    if (nameStart === 0 || !/\s/u.test(out[nameStart - 1] ?? "")) {
      searchFrom = hitIdx + SUFFIX.length;
      continue;
    }
    const quoteIdx = hitIdx + SUFFIX.length;
    const quote = out[quoteIdx];
    if (quote !== '"' && quote !== "'") {
      searchFrom = hitIdx + SUFFIX.length;
      continue;
    }
    const closeIdx = out.indexOf(quote, quoteIdx + 1);
    if (closeIdx === -1) {
      break;
    }
    // Drop the leading separator too so two whitespace runs do not merge.
    out = `${out.slice(0, nameStart - 1)}${out.slice(closeIdx + 1)}`;
    searchFrom = nameStart - 1;
  }
  return out;
}

const withCheckedState = (xml: string, checked: boolean): string => {
  const val = checked ? "1" : "0";
  // The expanded-empty form first, so the whole element is replaced rather
  // than only its start tag, which would leave a stray closing one behind.
  const opened = /<(?<prefix>\w+):checked\b[^>]*>[\s\S]*?<\/\w+:checked>/giu;
  if (/<\w+:checked\b[^>]*>[\s\S]*?<\/\w+:checked>/iu.test(xml)) {
    return xml.replaceAll(
      opened,
      (_match: string, prefix: string) => `<${prefix}:checked ${prefix}:val="${val}"/>`,
    );
  }
  if (/<\w+:checked\b[^>]*\/>/iu.test(xml)) {
    return xml.replaceAll(
      /<(?<prefix>\w+):checked\b[^>]*\/>/giu,
      (_match: string, prefix: string) => `<${prefix}:checked ${prefix}:val="${val}"/>`,
    );
  }
  // A `w14:checkbox` that states no checked child at all: fold one in under
  // the wrapper's own prefix, so the element stays in one namespace.
  return xml.replace(
    /<(?<prefix>\w+):checkbox\b[^>]*>/iu,
    (match: string, prefix: string) => `${match}<${prefix}:checked ${prefix}:val="${val}"/>`,
  );
};

const withDateState = (xml: string, props: SdtProperties): string => {
  const fullDate = props.dateValueISO;
  const format = props.dateFormat;
  if (fullDate === undefined && format === undefined) {
    return xml;
  }
  const fullDateAttr =
    fullDate === undefined ? "" : ` w:fullDate="${escapeXmlAttribute(fullDate)}"`;
  const formatChild =
    format === undefined ? "" : `<w:dateFormat w:val="${escapeXmlAttribute(format)}"/>`;
  const opened = /<(?<prefix>\w+):date\b(?<attrs>[^>]*)>(?<inner>[\s\S]*?)<\/\w+:date>/iu;
  if (opened.test(xml)) {
    return xml.replace(opened, (_match, prefix: string, matchedAttrs: string, inner: string) => {
      // One alternation covers both the self-closing and the expanded-empty
      // spelling of `w:dateFormat`; stripping only one would leave a stale
      // sibling beside the replacement on the next save.
      let body = inner.replaceAll(
        /<\w+:dateFormat\b[^>]*(?:\/>|>[\s\S]*?<\/\w+:dateFormat>)/giu,
        "",
      );
      if (formatChild) {
        body = `${formatChild}${body}`;
      }
      return `<${prefix}:date${fullDate === undefined ? matchedAttrs : fullDateAttr}>${body}</${prefix}:date>`;
    });
  }
  return xml.replace(
    /<(?<prefix>\w+):date\b(?<attrs>[^/>]*)\/>/iu,
    (_match, prefix: string) => `<${prefix}:date${fullDateAttr}>${formatChild}</${prefix}:date>`,
  );
};

const withLastValue = (xml: string, lastValue: string): string => {
  const escaped = escapeXmlAttribute(lastValue);
  const opened =
    /<(?<prefix>\w+):(?<name>dropDownList|comboBox)\b(?<attrs>[^>]*)>(?<inner>[\s\S]*?)<\/\w+:(?:dropDownList|comboBox)>/iu;
  if (opened.test(xml)) {
    return xml.replace(
      opened,
      (_match, prefix: string, name: string, attrs: string, inner: string) =>
        // Re-emitted under the SOURCE prefix, so a producer that bound the
        // Word namespace to `ns0` does not end up mixing prefixes inside one
        // element.
        `<${prefix}:${name}${stripLastValueAttr(attrs)} ${prefix}:lastValue="${escaped}">${inner}</${prefix}:${name}>`,
    );
  }
  return xml.replace(
    /<(?<prefix>\w+):(?<name>dropDownList|comboBox)\b(?<attrs>[^/>]*)\/>/iu,
    (_match, prefix: string, name: string, attrs: string) =>
      `<${prefix}:${name}${stripLastValueAttr(attrs)} ${prefix}:lastValue="${escaped}"/>`,
  );
};

/**
 * The control-kind element with the modelled interactive state written in.
 *
 * Pure: returns a new string; the input is not modified. A kind the model
 * states nothing interactive about comes back unchanged.
 */
export const withModelledControlState = (xml: string, props: SdtProperties): string => {
  switch (props.sdtType) {
    case "checkbox":
      return typeof props.checked === "boolean" ? withCheckedState(xml, props.checked) : xml;
    case "date":
      return withDateState(xml, props);
    case "dropdown":
    case "comboBox":
      return props.dropdownLastValue === undefined
        ? xml
        : withLastValue(xml, props.dropdownLastValue);
    case "richText":
    case "plainText":
    case "picture":
    case "buildingBlockGallery":
    case "group":
    case "unknown":
      return xml;
    default: {
      const exhaustive: never = props.sdtType;
      return exhaustive;
    }
  }
};
