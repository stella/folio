import { XMLValidator } from "fast-xml-parser";

const DOCTYPE_PATTERN = /<!DOCTYPE(?:\s|>)/iu;
const BUILT_IN_ENTITIES = new Set(["amp", "apos", "gt", "lt", "quot"]);
const DECIMAL_CHARACTER_REFERENCE = /^#[0-9]+$/u;
const HEX_CHARACTER_REFERENCE = /^#[xX][0-9A-Fa-f]+$/u;
const NCNAME = /^[\p{L}_][\p{L}\p{N}._\-\u00b7\p{M}]*$/u;
const BYTE_ORDER_MARK = "\uFEFF";
const ENCODING_NAME = String.raw`[A-Za-z][A-Za-z0-9._-]*`;
const XML_DECLARATION = new RegExp(
  String.raw`^<\?xml[ \t\r\n]+version[ \t\r\n]*=[ \t\r\n]*(?:"1\.0"|'1\.0')(?:[ \t\r\n]+encoding[ \t\r\n]*=[ \t\r\n]*(?:"${ENCODING_NAME}"|'${ENCODING_NAME}'))?(?:[ \t\r\n]+standalone[ \t\r\n]*=[ \t\r\n]*(?:"(?:yes|no)"|'(?:yes|no)'))?[ \t\r\n]*\?>$`,
  "u",
);

export type DocxXmlSafetyIssue = "doctype-forbidden" | "entity-forbidden" | "not-well-formed";

const isXmlCharacter = (codePoint: number): boolean =>
  codePoint === 0x09 ||
  codePoint === 0x0a ||
  codePoint === 0x0d ||
  (codePoint >= 0x20 && codePoint <= 0xd7ff) ||
  (codePoint >= 0xe000 && codePoint <= 0xfffd) ||
  (codePoint >= 0x10000 && codePoint <= 0x10ffff);

const validEntity = (entity: string): boolean => {
  if (BUILT_IN_ENTITIES.has(entity)) {
    return true;
  }
  if (!DECIMAL_CHARACTER_REFERENCE.test(entity) && !HEX_CHARACTER_REFERENCE.test(entity)) {
    return false;
  }
  const digits = entity[1]?.toLowerCase() === "x" ? entity.slice(2) : entity.slice(1);
  const significantDigits = digits.replace(/^0+/u, "") || "0";
  const hexadecimal = entity[1]?.toLowerCase() === "x";
  if (digits.length === 0 || significantDigits.length > (hexadecimal ? 6 : 7)) {
    return false;
  }
  return isXmlCharacter(Number.parseInt(significantDigits, hexadecimal ? 16 : 10));
};

const hasForbiddenEntity = (xml: string): boolean => {
  let cursor = 0;
  while (cursor < xml.length) {
    if (xml.startsWith("<!--", cursor)) {
      const end = xml.indexOf("-->", cursor + 4);
      cursor = end === -1 ? xml.length : end + 3;
      continue;
    }
    if (xml.startsWith("<![CDATA[", cursor)) {
      const end = xml.indexOf("]]>", cursor + 9);
      cursor = end === -1 ? xml.length : end + 3;
      continue;
    }
    if (xml.startsWith("<?", cursor)) {
      const end = xml.indexOf("?>", cursor + 2);
      cursor = end === -1 ? xml.length : end + 2;
      continue;
    }
    if (xml[cursor] !== "&") {
      cursor += 1;
      continue;
    }
    const end = xml.indexOf(";", cursor + 1);
    if (end === -1) {
      return true;
    }
    if (!validEntity(xml.slice(cursor + 1, end))) {
      return true;
    }
    cursor = end + 1;
  }
  return false;
};

const isValidQName = (name: string): boolean => {
  const parts = name.split(":");
  return parts.length <= 2 && parts.every((part) => NCNAME.test(part));
};

const containsOnlyXmlCharacters = (xml: string): boolean => {
  for (const character of xml) {
    const codePoint = character.codePointAt(0);
    if (codePoint === undefined || !isXmlCharacter(codePoint)) {
      return false;
    }
  }
  return true;
};

const isWhitespace = (character: string | undefined): boolean =>
  character === " " || character === "\t" || character === "\r" || character === "\n";

type ScannedName = { name: string; next: number };

const scanName = (xml: string, start: number): ScannedName | null => {
  let cursor = start;
  while (
    cursor < xml.length &&
    !isWhitespace(xml[cursor]) &&
    xml[cursor] !== ">" &&
    xml[cursor] !== "/" &&
    xml[cursor] !== "=" &&
    xml[cursor] !== "?"
  ) {
    cursor += 1;
  }
  const name = xml.slice(start, cursor);
  return name !== "" && isValidQName(name) ? { name, next: cursor } : null;
};

/**
 * Reject lexical forms that fast-xml-parser's permissive validator repairs.
 * This is deliberately a small XML lexer, not a second tree parser: structural
 * nesting remains XMLValidator's job.
 */
const hasInvalidXmlLexicalForm = (xml: string): boolean => {
  if (!containsOnlyXmlCharacters(xml)) {
    return true;
  }
  const byteOrderMark = xml.indexOf(BYTE_ORDER_MARK);
  if (
    byteOrderMark > 0 ||
    (byteOrderMark === 0 && xml.indexOf(BYTE_ORDER_MARK, BYTE_ORDER_MARK.length) !== -1)
  ) {
    return true;
  }

  let cursor = 0;
  while (cursor < xml.length) {
    const opening = xml.indexOf("<", cursor);
    const textEnd = opening === -1 ? xml.length : opening;
    if (xml.slice(cursor, textEnd).includes("]]>") || opening === -1) {
      return xml.slice(cursor, textEnd).includes("]]>");
    }

    if (xml.startsWith("<!--", opening)) {
      const end = xml.indexOf("-->", opening + 4);
      if (end === -1) {
        return true;
      }
      const body = xml.slice(opening + 4, end);
      if (body.includes("--") || body.endsWith("-")) {
        return true;
      }
      cursor = end + 3;
      continue;
    }

    if (xml.startsWith("<![CDATA[", opening)) {
      const end = xml.indexOf("]]>", opening + 9);
      if (end === -1) {
        return true;
      }
      cursor = end + 3;
      continue;
    }

    if (xml.startsWith("<?", opening)) {
      const end = xml.indexOf("?>", opening + 2);
      if (end === -1) {
        return true;
      }
      let targetEnd = opening + 2;
      while (targetEnd < end && !isWhitespace(xml[targetEnd]) && xml[targetEnd] !== "?") {
        targetEnd += 1;
      }
      const target = xml.slice(opening + 2, targetEnd);
      if (!NCNAME.test(target)) {
        return true;
      }
      if (target.toLowerCase() === "xml") {
        // The XML declaration is the sole reserved `xml` processing target.
        const declarationOpening = xml.startsWith(BYTE_ORDER_MARK) ? BYTE_ORDER_MARK.length : 0;
        if (
          opening !== declarationOpening ||
          target !== "xml" ||
          !XML_DECLARATION.test(xml.slice(opening, end + 2))
        ) {
          return true;
        }
      }
      cursor = end + 2;
      continue;
    }

    if (xml.startsWith("<!", opening)) {
      return true;
    }

    let tagCursor = opening + 1;
    const closing = xml[tagCursor] === "/";
    if (closing) {
      tagCursor += 1;
    }
    const tagName = scanName(xml, tagCursor);
    if (!tagName) {
      return true;
    }
    tagCursor = tagName.next;

    if (closing) {
      while (isWhitespace(xml[tagCursor])) {
        tagCursor += 1;
      }
      if (xml[tagCursor] !== ">") {
        return true;
      }
      cursor = tagCursor + 1;
      continue;
    }

    while (tagCursor < xml.length) {
      while (isWhitespace(xml[tagCursor])) {
        tagCursor += 1;
      }
      if (xml[tagCursor] === ">") {
        cursor = tagCursor + 1;
        break;
      }
      if (xml[tagCursor] === "/" && xml[tagCursor + 1] === ">") {
        cursor = tagCursor + 2;
        break;
      }

      const attribute = scanName(xml, tagCursor);
      if (!attribute) {
        return true;
      }
      if (
        attribute.name.startsWith("xmlns:") &&
        (!NCNAME.test(attribute.name.slice("xmlns:".length)) ||
          attribute.name.slice("xmlns:".length).includes(":"))
      ) {
        return true;
      }
      tagCursor = attribute.next;
      while (isWhitespace(xml[tagCursor])) {
        tagCursor += 1;
      }
      if (xml[tagCursor] !== "=") {
        return true;
      }
      tagCursor += 1;
      while (isWhitespace(xml[tagCursor])) {
        tagCursor += 1;
      }
      const quote = xml[tagCursor];
      if (quote !== '"' && quote !== "'") {
        return true;
      }
      const valueEnd = xml.indexOf(quote, tagCursor + 1);
      if (valueEnd === -1 || xml.slice(tagCursor + 1, valueEnd).includes("<")) {
        return true;
      }
      tagCursor = valueEnd + 1;
    }
    if (tagCursor >= xml.length && cursor !== xml.length) {
      return true;
    }
  }
  return false;
};

export const getDocxXmlSafetyIssue = (xml: string): DocxXmlSafetyIssue | null => {
  if (DOCTYPE_PATTERN.test(xml)) {
    return "doctype-forbidden";
  }
  if (hasForbiddenEntity(xml)) {
    return "entity-forbidden";
  }
  if (hasInvalidXmlLexicalForm(xml)) {
    return "not-well-formed";
  }
  return XMLValidator.validate(xml) === true ? null : "not-well-formed";
};
