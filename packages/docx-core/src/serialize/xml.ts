import { escapeXmlAttribute } from "./xmlEscape";

export const attr = (name: string, value: string | number | boolean | undefined) =>
  value === undefined ? "" : ` ${name}="${escapeXmlAttribute(String(value))}"`;
