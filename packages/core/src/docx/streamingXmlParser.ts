import type { XmlElement } from "./xmlParser";

type ParseXmlResult = { status: "parsed"; value: XmlElement } | { status: "unsupported" };

type ElementFrame = {
  element: XmlElement;
  name: string;
};

const BUILT_IN_ENTITIES = {
  amp: "&",
  apos: "'",
  gt: ">",
  lt: "<",
  quot: '"',
} as const;

/**
 * Parse ordinary OOXML into Folio's existing element representation in one
 * pass. Unsupported or malformed constructs return a sentinel so callers can
 * retain the general-purpose parser as a compatibility fallback.
 */
export const parseStreamingXml = (xml: string): ParseXmlResult => {
  const root: XmlElement = { elements: [] };
  const stack: ElementFrame[] = [];
  let cursor = 0;
  let mergeAdjacentText = false;

  while (cursor < xml.length) {
    const open = xml.indexOf("<", cursor);
    if (open === -1) {
      if (!appendText(xml.slice(cursor), stack, mergeAdjacentText)) {
        return { status: "unsupported" };
      }
      break;
    }

    if (open > cursor && !appendText(xml.slice(cursor, open), stack, mergeAdjacentText)) {
      return { status: "unsupported" };
    }
    if (open > cursor) {
      mergeAdjacentText = true;
    }

    if (xml.startsWith("<!--", open)) {
      const close = xml.indexOf("-->", open + 4);
      if (close === -1) {
        return { status: "unsupported" };
      }
      cursor = close + 3;
      continue;
    }

    if (xml.startsWith("<![CDATA[", open)) {
      const close = xml.indexOf("]]>", open + 9);
      if (close === -1 || !appendRawText(xml.slice(open + 9, close), stack)) {
        return { status: "unsupported" };
      }
      cursor = close + 3;
      mergeAdjacentText = false;
      continue;
    }

    if (xml.startsWith("<?", open)) {
      const close = xml.indexOf("?>", open + 2);
      if (close === -1) {
        return { status: "unsupported" };
      }
      cursor = close + 2;
      continue;
    }

    if (xml.startsWith("<!", open)) {
      return { status: "unsupported" };
    }

    const close = findTagClose(xml, open + 1);
    if (close === -1) {
      return { status: "unsupported" };
    }

    if (xml.charCodeAt(open + 1) === 47) {
      const name = xml.slice(open + 2, close).trim();
      const frame = stack.pop();
      if (!frame || frame.name !== name) {
        return { status: "unsupported" };
      }
      cursor = close + 1;
      mergeAdjacentText = false;
      continue;
    }

    const parsedTag = parseOpenTag(xml, open + 1, close);
    if (parsedTag.status === "unsupported") {
      return parsedTag;
    }

    const parent = stack.at(-1)?.element ?? root;
    appendElement(parent, parsedTag.element);
    if (!parsedTag.selfClosing) {
      if (stack.length >= 100) {
        return { status: "unsupported" };
      }
      stack.push({ element: parsedTag.element, name: parsedTag.name });
    }
    cursor = close + 1;
    mergeAdjacentText = false;
  }

  if (stack.length > 0) {
    return { status: "unsupported" };
  }
  return { status: "parsed", value: root };
};

type ParsedOpenTag =
  | {
      status: "parsed";
      element: XmlElement;
      name: string;
      selfClosing: boolean;
    }
  | { status: "unsupported" };

const parseOpenTag = (xml: string, start: number, close: number): ParsedOpenTag => {
  let cursor = skipWhitespace(xml, start, close);
  const nameStart = cursor;
  cursor = scanName(xml, cursor, close);
  if (cursor === nameStart) {
    return { status: "unsupported" };
  }

  const name = xml.slice(nameStart, cursor);
  if (isUnsafePropertyName(name)) {
    return { status: "unsupported" };
  }
  let attributes: Record<string, string> | undefined;
  let selfClosing = false;

  while (cursor < close) {
    cursor = skipWhitespace(xml, cursor, close);
    if (cursor >= close) {
      break;
    }
    if (xml.charCodeAt(cursor) === 47) {
      selfClosing = true;
      cursor = skipWhitespace(xml, cursor + 1, close);
      if (cursor !== close) {
        return { status: "unsupported" };
      }
      break;
    }

    const attributeNameStart = cursor;
    cursor = scanName(xml, cursor, close);
    if (cursor === attributeNameStart) {
      return { status: "unsupported" };
    }
    const attributeName = xml.slice(attributeNameStart, cursor);
    if (isUnsafePropertyName(attributeName)) {
      return { status: "unsupported" };
    }
    cursor = skipWhitespace(xml, cursor, close);
    if (xml.charCodeAt(cursor) !== 61) {
      return { status: "unsupported" };
    }
    cursor = skipWhitespace(xml, cursor + 1, close);

    const quote = xml.charCodeAt(cursor);
    if (quote !== 34 && quote !== 39) {
      return { status: "unsupported" };
    }
    const valueStart = cursor + 1;
    cursor = xml.indexOf(quote === 34 ? '"' : "'", valueStart);
    if (cursor === -1 || cursor > close) {
      return { status: "unsupported" };
    }
    const decoded = decodeXmlEntities(normalizeLineEndings(xml.slice(valueStart, cursor)));
    if (decoded === null) {
      return { status: "unsupported" };
    }
    attributes ??= {};
    attributes[attributeName] = decoded;
    cursor += 1;
  }

  const element: XmlElement = { type: "element", name };
  if (attributes) {
    element.attributes = attributes;
  }
  return { status: "parsed", element, name, selfClosing };
};

const appendElement = (parent: XmlElement, child: XmlElement): void => {
  parent.elements ??= [];
  parent.elements.push(child);
};

const appendText = (text: string, stack: ElementFrame[], mergeAdjacentText: boolean): boolean => {
  if (text.length === 0) {
    return true;
  }
  const decoded = decodeXmlEntities(normalizeLineEndings(text));
  if (decoded === null) {
    return false;
  }
  const parent = stack.at(-1)?.element;
  const previous = mergeAdjacentText ? parent?.elements?.at(-1) : undefined;
  if (previous?.type === "text" && typeof previous.text === "string") {
    previous.text += decoded;
    return true;
  }
  return appendRawText(decoded, stack);
};

const appendRawText = (text: string, stack: ElementFrame[]): boolean => {
  const parent = stack.at(-1)?.element;
  if (!parent) {
    return text.trim().length === 0;
  }
  appendElement(parent, { type: "text", text });
  return true;
};

const findTagClose = (xml: string, start: number): number => {
  let quote = 0;
  for (let cursor = start; cursor < xml.length; cursor += 1) {
    const code = xml.charCodeAt(cursor);
    if (quote !== 0) {
      if (code === quote) {
        quote = 0;
      }
      continue;
    }
    if (code === 34 || code === 39) {
      quote = code;
      continue;
    }
    if (code === 62) {
      return cursor;
    }
  }
  return -1;
};

const skipWhitespace = (xml: string, start: number, limit: number): number => {
  let cursor = start;
  while (cursor < limit) {
    const code = xml.charCodeAt(cursor);
    if (code !== 9 && code !== 10 && code !== 13 && code !== 32) {
      break;
    }
    cursor += 1;
  }
  return cursor;
};

const scanName = (xml: string, start: number, limit: number): number => {
  let cursor = start;
  while (cursor < limit) {
    const code = xml.charCodeAt(cursor);
    if (
      code === 9 ||
      code === 10 ||
      code === 13 ||
      code === 32 ||
      code === 47 ||
      code === 61 ||
      code === 62
    ) {
      break;
    }
    cursor += 1;
  }
  return cursor;
};

const decodeXmlEntities = (value: string): string | null => {
  const firstEntity = value.indexOf("&");
  if (firstEntity === -1) {
    return value;
  }

  const parts: string[] = [];
  let cursor = 0;
  let entityStart = firstEntity;
  while (entityStart !== -1) {
    parts.push(value.slice(cursor, entityStart));
    const entityEnd = value.indexOf(";", entityStart + 1);
    if (entityEnd === -1) {
      return null;
    }
    const entity = value.slice(entityStart + 1, entityEnd);
    const decoded = decodeXmlEntity(entity);
    if (decoded === null) {
      return null;
    }
    parts.push(decoded);
    cursor = entityEnd + 1;
    entityStart = value.indexOf("&", cursor);
  }
  parts.push(value.slice(cursor));
  return parts.join("");
};

const normalizeLineEndings = (value: string): string => {
  if (!value.includes("\r")) {
    return value;
  }
  return value.replaceAll("\r\n", "\n").replaceAll("\r", "\n");
};

const decodeXmlEntity = (entity: string): string | null => {
  switch (entity) {
    case "amp":
      return BUILT_IN_ENTITIES.amp;
    case "apos":
      return BUILT_IN_ENTITIES.apos;
    case "gt":
      return BUILT_IN_ENTITIES.gt;
    case "lt":
      return BUILT_IN_ENTITIES.lt;
    case "quot":
      return BUILT_IN_ENTITIES.quot;
  }
  if (entity.charCodeAt(0) !== 35) {
    return null;
  }

  const hexadecimal = entity.charCodeAt(1) === 120 || entity.charCodeAt(1) === 88;
  const digits = entity.slice(hexadecimal ? 2 : 1);
  if (digits.length === 0 || !hasOnlyDigits(digits, hexadecimal)) {
    return null;
  }
  const codePoint = Number.parseInt(digits, hexadecimal ? 16 : 10);
  if (
    !Number.isInteger(codePoint) ||
    codePoint <= 0 ||
    codePoint > 0x10ffff ||
    (codePoint >= 0xd800 && codePoint <= 0xdfff)
  ) {
    return null;
  }
  return String.fromCodePoint(codePoint);
};

const hasOnlyDigits = (value: string, hexadecimal: boolean): boolean => {
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index);
    const decimal = code >= 48 && code <= 57;
    const lowerHex = hexadecimal && code >= 97 && code <= 102;
    const upperHex = hexadecimal && code >= 65 && code <= 70;
    if (!decimal && !lowerHex && !upperHex) {
      return false;
    }
  }
  return true;
};

const isUnsafePropertyName = (name: string): boolean => {
  switch (name) {
    case "__defineGetter__":
    case "__defineSetter__":
    case "__lookupGetter__":
    case "__lookupSetter__":
    case "__proto__":
    case "constructor":
    case "hasOwnProperty":
    case "prototype":
    case "toString":
    case "valueOf":
      return true;
    default:
      return false;
  }
};
