import type { RunContent } from "../types/document";
import {
  findChildByNamespaceUri,
  getAttributeByNamespaceUri,
  NAMESPACES,
  OOXML_NAMESPACE_SCOPE,
  parseXml,
  WORDPROCESSINGML_NAMESPACE_URIS,
} from "../docx/xmlParser";

const CHECKBOX_NAMESPACE_URIS: ReadonlySet<string> = new Set([
  NAMESPACES.w14,
  ...WORDPROCESSINGML_NAMESPACE_URIS,
]);
const DEFAULT_CHECKED_CHARACTER = "2612";
const DEFAULT_UNCHECKED_CHARACTER = "2610";
const DEFAULT_CHECKBOX_FONT = "MS Gothic";
const HEX_CODE_POINT = /^[\dA-Fa-f]{1,6}$/u;
const MAX_CHECKBOX_PROPERTIES_CHARACTERS = 65_536;

type CheckboxDisplayContent = Extract<RunContent, { type: "symbol" } | { type: "text" }>;

const defaultDisplay = (checked: boolean): CheckboxDisplayContent => ({
  type: "symbol",
  char: checked ? DEFAULT_CHECKED_CHARACTER : DEFAULT_UNCHECKED_CHARACTER,
  font: DEFAULT_CHECKBOX_FONT,
});

/**
 * Resolve the authored checkbox state into the run payload Word expects.
 *
 * @param checkboxXml the preserved `w14:checkbox` element, from
 *   `SdtProperties.preserved`. A control that carries none — one built in
 *   code, or one whose author wrote no glyphs — gets Word's own defaults.
 */
export const checkboxDisplayContent = (
  checkboxXml: string | undefined,
  checked: boolean,
): CheckboxDisplayContent => {
  if (!checkboxXml || checkboxXml.length > MAX_CHECKBOX_PROPERTIES_CHARACTERS) {
    return defaultDisplay(checked);
  }

  let root;
  try {
    root = parseXml(checkboxXml, OOXML_NAMESPACE_SCOPE);
  } catch {
    return defaultDisplay(checked);
  }
  const checkbox = findChildByNamespaceUri(root, CHECKBOX_NAMESPACE_URIS, "checkbox");
  const state = findChildByNamespaceUri(
    checkbox,
    CHECKBOX_NAMESPACE_URIS,
    checked ? "checkedState" : "uncheckedState",
  );
  const rawCharacter = getAttributeByNamespaceUri(state, CHECKBOX_NAMESPACE_URIS, "val");
  if (!rawCharacter || !HEX_CODE_POINT.test(rawCharacter)) {
    return defaultDisplay(checked);
  }

  const codePoint = Number.parseInt(rawCharacter, 16);
  if (
    !Number.isInteger(codePoint) ||
    codePoint > 0x10ffff ||
    (codePoint >= 0xd800 && codePoint <= 0xdfff)
  ) {
    return defaultDisplay(checked);
  }
  const font =
    getAttributeByNamespaceUri(state, CHECKBOX_NAMESPACE_URIS, "font") ?? DEFAULT_CHECKBOX_FONT;
  if (codePoint <= 0xffff) {
    return { type: "symbol", char: codePoint.toString(16).toUpperCase().padStart(4, "0"), font };
  }
  return { type: "text", text: String.fromCodePoint(codePoint) };
};
