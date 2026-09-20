import { escapeXmlAttribute } from "./xmlEscape";

export const attr = (name: string, value: string | number | boolean | undefined) =>
  value === undefined ? "" : ` ${name}="${escapeXmlAttribute(String(value))}"`;

/**
 * The one writer for a `CT_OnOff` element: `w:b`, `w:cantSplit`, `w:hideMark`,
 * every element the schema types `CT_OnOff`.
 *
 * A toggle has three states, not two. Absent inherits whatever the style
 * hierarchy resolves to; the bare element is an explicit on; `w:val="0"` is an
 * explicit off, and an explicit off is the only thing that cancels an inherited
 * on. ECMA-376 combines two levels of the hierarchy by XOR rather than by
 * override, so dropping the off does not merely lose a `false`: it flips the
 * resolved value of every toggle the style set above it turns on.
 *
 * Writing the bare element for on and nothing at all for off — the shape a
 * `if (value) parts.push("<w:x/>")` guard produces — is therefore a silent
 * corruption, and it is a shape that reappears at every new toggle site. So
 * there is one writer, `scripts/on-off-element-writer.test.ts` refuses a second
 * one, and a site that spells the element itself does not reach a consumer.
 *
 * `ST_OnOff` spells an on three ways and an off three ways. The readers take
 * all six, and a captured element replays the bytes it arrived as; what folio
 * *writes* is one spelling per polarity, because a package that writes three
 * makes every byte-level comparison argue about which one it is looking at.
 */
export const serializeOnOffElement = (value: boolean | undefined, name: string): string => {
  if (value === undefined) {
    return "";
  }
  return value ? `<w:${name}/>` : `<w:${name} w:val="0"/>`;
};

/**
 * Append a `CT_OnOff` element to a property list, and nothing when it is absent.
 *
 * Most property serializers decide whether to write the enclosing element from
 * `parts.length`, so an absent toggle must not leave an empty string behind:
 * that would turn an unformatted run into a `<w:rPr></w:rPr>`.
 */
export const pushOnOffElement = (
  parts: string[],
  value: boolean | undefined,
  name: string,
): void => {
  const xml = serializeOnOffElement(value, name);
  if (xml) {
    parts.push(xml);
  }
};
