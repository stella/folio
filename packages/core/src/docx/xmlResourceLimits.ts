import { TaggedError } from "better-result";

const MEBIBYTE = 1024 * 1024;

/** Shared bounds for XML parts parsed by Folio. */
export const FOLIO_XML_RESOURCE_LIMITS = {
  maxBytes: 128 * MEBIBYTE,
  maxDepth: 100,
  maxNodes: 1_000_000,
} as const;

type XmlResourceLimits = {
  maxBytes: number;
  maxDepth: number;
  maxNodes: number;
};

type XmlResourceLimitKind = "bytes" | "depth" | "nodes" | "syntax";

/** XML input exceeded a parser resource bound or could not be scanned safely. */
export class XmlResourceLimitError extends TaggedError("XmlResourceLimitError")<{
  message: string;
  limit: XmlResourceLimitKind;
}> {}

const exceedsUtf8ByteLimit = (value: string, maxBytes: number): boolean => {
  let bytes = 0;
  for (let index = 0; index < value.length; index += 1) {
    const codeUnit = value.charCodeAt(index);
    if (codeUnit <= 0x7f) {
      bytes += 1;
    } else if (codeUnit <= 0x7ff) {
      bytes += 2;
    } else if (
      codeUnit >= 0xd800 &&
      codeUnit <= 0xdbff &&
      index + 1 < value.length &&
      value.charCodeAt(index + 1) >= 0xdc00 &&
      value.charCodeAt(index + 1) <= 0xdfff
    ) {
      bytes += 4;
      index += 1;
    } else {
      bytes += 3;
    }
    if (bytes > maxBytes) {
      return true;
    }
  }
  return false;
};

const isXmlWhitespace = (code: number): boolean =>
  code === 9 || code === 10 || code === 13 || code === 32;

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

const throwSyntaxLimit = (): never => {
  throw new XmlResourceLimitError({
    message: "XML resource preflight could not safely scan malformed markup",
    limit: "syntax",
  });
};

/**
 * Bound XML bytes, element count, and nesting before building an object tree.
 * The lexical scan is iterative, so deeply nested input cannot consume the JS
 * call stack before the depth limit is enforced.
 */
export const assertXmlResourceLimits = (
  xml: string,
  limits: XmlResourceLimits = FOLIO_XML_RESOURCE_LIMITS,
): void => {
  if (exceedsUtf8ByteLimit(xml, limits.maxBytes)) {
    throw new XmlResourceLimitError({
      message: `XML part exceeds ${String(limits.maxBytes)} bytes`,
      limit: "bytes",
    });
  }

  let cursor = 0;
  let depth = 0;
  let nodes = 0;
  while (cursor < xml.length) {
    const open = xml.indexOf("<", cursor);
    if (open === -1) {
      break;
    }
    if (xml.startsWith("<!--", open)) {
      const close = xml.indexOf("-->", open + 4);
      if (close === -1) {
        throwSyntaxLimit();
      }
      cursor = close + 3;
      continue;
    }
    if (xml.startsWith("<![CDATA[", open)) {
      const close = xml.indexOf("]]>", open + 9);
      if (close === -1) {
        throwSyntaxLimit();
      }
      cursor = close + 3;
      continue;
    }
    if (xml.startsWith("<?", open)) {
      const close = xml.indexOf("?>", open + 2);
      if (close === -1) {
        throwSyntaxLimit();
      }
      cursor = close + 2;
      continue;
    }
    if (xml.startsWith("<!", open)) {
      throwSyntaxLimit();
    }

    const close = findTagClose(xml, open + 1);
    if (close === -1) {
      throwSyntaxLimit();
    }
    if (xml.charCodeAt(open + 1) === 47) {
      if (depth === 0) {
        throwSyntaxLimit();
      }
      depth -= 1;
      cursor = close + 1;
      continue;
    }

    nodes += 1;
    if (nodes > limits.maxNodes) {
      throw new XmlResourceLimitError({
        message: `XML part contains more than ${String(limits.maxNodes)} elements`,
        limit: "nodes",
      });
    }

    let lastContent = close - 1;
    while (lastContent > open && isXmlWhitespace(xml.charCodeAt(lastContent))) {
      lastContent -= 1;
    }
    const elementDepth = depth + 1;
    if (elementDepth > limits.maxDepth) {
      throw new XmlResourceLimitError({
        message: `XML part is nested deeper than ${String(limits.maxDepth)} elements`,
        limit: "depth",
      });
    }
    if (xml.charCodeAt(lastContent) !== 47) {
      depth = elementDepth;
    }
    cursor = close + 1;
  }

  if (depth !== 0) {
    throwSyntaxLimit();
  }
};
