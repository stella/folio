import { TaggedError } from "better-result";

const MEBIBYTE = 1024 * 1024;

/**
 * Shared bounds for XML parts parsed by Folio.
 *
 * `maxBytes` and the package expansion ceiling bound the *markup*; they do not
 * bound what parsing that markup allocates. A parsed element retains far more
 * than the bytes it was written as, so a part that satisfies a byte bound can
 * still cost multiples of it in tree. Measured on this repository's corpus
 * generators (Bun 1.4, arm64, heap retained after a forced GC):
 *
 *   shape                       bytes/element   tree heap / part bytes
 *   `<w:r/>`                          49.5 B            8.3x
 *   `<w:r a=".." x4/>`               105-124 B          2.8-3.3x
 *   `<w:r><w:t>x</w:t></w:r>`        162-171 B         14.1-14.8x
 *
 * An element is therefore the unit a memory bound has to count, because the
 * adversarial shape is the cheap one: `<w:r/>` is 7 bytes of markup and 49.5
 * bytes of tree, so 128 MiB of markup buys ~19M elements and ~950 MB of tree
 * inside a byte budget that never trips.
 *
 * Corpus distribution (5,314 readable packages, every XML and .rels part):
 *
 *   metric                     p50      p90      p99    p99.9        max
 *   elements / part             16      227    1,590   24,719    596,668
 *   elements / package         917    3,230   29,106   71,975    602,212
 *   attributes / part           27      577    1,993   21,037    630,374
 *   attributes / package     1,783    4,711   25,896   91,946    639,110
 *   depth / part                 3        9       15       22     15,005
 *   elements per byte        0.012    0.027    0.036    0.043      0.111
 *
 * The defaults below reject no corpus package that today's bounds accept. The
 * package budget is the one that matters: it caps a whole package at ~2.5M
 * elements and ~3M attributes, which is at most ~425 MB of tree at the densest
 * measured shape and ~124 MB at the cheapest. Before it existed, only
 * `word/document.xml`, `word/styles.xml` and `word/numbering.xml` were counted
 * at all, so a package of many merely-large parts could reach the expansion
 * ceiling of 250 MiB, ~37M elements and well past 1.8 GB of tree while passing
 * every bound. Lower these to trade format reach for a smaller ceiling.
 */
export const FOLIO_XML_RESOURCE_LIMITS = {
  maxBytes: 128 * MEBIBYTE,
  maxDepth: 100,
  /** 1.68x the corpus maximum (596,668). Unchanged; a shipped bound is not loosened. */
  maxElementsPerPart: 1_000_000,
  /** 3.97x the corpus maximum (630,374). */
  maxAttributesPerPart: 2_500_000,
  /** 4.15x the corpus maximum (602,212). */
  maxElementsPerPackage: 2_500_000,
  /** 4.69x the corpus maximum (639,110). */
  maxAttributesPerPackage: 3_000_000,
} as const;

export type XmlResourceLimits = {
  maxBytes: number;
  maxDepth: number;
  maxElementsPerPart: number;
  maxAttributesPerPart: number;
  maxElementsPerPackage: number;
  maxAttributesPerPackage: number;
};

type XmlResourceLimitKind =
  | "bytes"
  | "depth"
  | "elements"
  | "attributes"
  | "package-elements"
  | "package-attributes"
  | "syntax";

/** XML input exceeded a parser resource bound or could not be scanned safely. */
export class XmlResourceLimitError extends TaggedError("XmlResourceLimitError")<{
  message: string;
  limit: XmlResourceLimitKind;
  /** The package part being scanned, when the caller named one. */
  partPath?: string;
  /** The count reached at the point of refusal, not the count of the whole input. */
  observed: number;
  allowed: number;
}> {}

/**
 * What a package has spent so far, shared by every part in one package.
 *
 * Per-part bounds alone do not bound a package: a package may hold hundreds of
 * parts, each individually modest. The budget is the accumulator the readers
 * carry across parts so the ceiling is the package's, not each part's.
 */
export type XmlPackageBudget = {
  elements: number;
  attributes: number;
};

export const createXmlPackageBudget = (): XmlPackageBudget => ({ elements: 0, attributes: 0 });

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

/**
 * The tag's closing `>` and how many attributes it declares.
 *
 * Attributes are counted in the same pass that finds the close: outside a
 * quoted value an `=` in a tag is an attribute assignment and nothing else, so
 * the count is exact for markup the scan accepts, and free.
 */
type TagScan = { close: number; attributes: number };

const scanTag = (xml: string, start: number): TagScan => {
  let quote = 0;
  let attributes = 0;
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
    if (code === 61) {
      attributes += 1;
      continue;
    }
    if (code === 62) {
      return { close: cursor, attributes };
    }
  }
  return { close: -1, attributes };
};

const throwSyntaxLimit = (): never => {
  throw new XmlResourceLimitError({
    message: "XML resource preflight could not safely scan malformed markup",
    limit: "syntax",
    observed: 0,
    allowed: 0,
  });
};

export type XmlResourceScanOptions = {
  xml: string;
  limits?: XmlResourceLimits;
  /** The package path scanned, carried on any refusal so a host can name it. */
  partPath?: string;
  /** Charged as the scan proceeds; omit for a part with no package context. */
  budget?: XmlPackageBudget;
};

/** What the preflight counted, for callers that reconcile it against the tree. */
export type XmlResourceScanResult = {
  elements: number;
  attributes: number;
  maxDepth: number;
};

/**
 * Bound XML bytes, element count, attribute count and nesting before building
 * an object tree. The lexical scan is iterative, so deeply nested input cannot
 * consume the JS call stack before the depth limit is enforced, and every
 * bound is checked at the point it is crossed, so the work done before a
 * refusal is proportional to the limit rather than to the input.
 */
export const assertXmlResourceLimits = ({
  xml,
  limits = FOLIO_XML_RESOURCE_LIMITS,
  partPath,
  budget,
}: XmlResourceScanOptions): XmlResourceScanResult => {
  const refuse = (
    limit: XmlResourceLimitKind,
    message: string,
    observed: number,
    allowed: number,
  ): never => {
    throw new XmlResourceLimitError({
      message: partPath === undefined ? message : `${message} (part ${partPath})`,
      limit,
      ...(partPath === undefined ? {} : { partPath }),
      observed,
      allowed,
    });
  };

  if (exceedsUtf8ByteLimit(xml, limits.maxBytes)) {
    refuse(
      "bytes",
      `XML part exceeds ${String(limits.maxBytes)} bytes`,
      limits.maxBytes + 1,
      limits.maxBytes,
    );
  }

  const budgetElements = budget?.elements ?? 0;
  const budgetAttributes = budget?.attributes ?? 0;
  let cursor = 0;
  let depth = 0;
  let maxDepth = 0;
  let nodes = 0;
  let attributes = 0;
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

    const { close, attributes: tagAttributes } = scanTag(xml, open + 1);
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
    if (nodes > limits.maxElementsPerPart) {
      refuse(
        "elements",
        `XML part contains more than ${String(limits.maxElementsPerPart)} elements`,
        nodes,
        limits.maxElementsPerPart,
      );
    }
    if (budgetElements + nodes > limits.maxElementsPerPackage) {
      refuse(
        "package-elements",
        `DOCX package contains more than ${String(limits.maxElementsPerPackage)} XML elements`,
        budgetElements + nodes,
        limits.maxElementsPerPackage,
      );
    }

    attributes += tagAttributes;
    if (attributes > limits.maxAttributesPerPart) {
      refuse(
        "attributes",
        `XML part contains more than ${String(limits.maxAttributesPerPart)} attributes`,
        attributes,
        limits.maxAttributesPerPart,
      );
    }
    if (budgetAttributes + attributes > limits.maxAttributesPerPackage) {
      refuse(
        "package-attributes",
        `DOCX package contains more than ${String(limits.maxAttributesPerPackage)} XML attributes`,
        budgetAttributes + attributes,
        limits.maxAttributesPerPackage,
      );
    }

    let lastContent = close - 1;
    while (lastContent > open && isXmlWhitespace(xml.charCodeAt(lastContent))) {
      lastContent -= 1;
    }
    const elementDepth = depth + 1;
    if (elementDepth > limits.maxDepth) {
      refuse(
        "depth",
        `XML part is nested deeper than ${String(limits.maxDepth)} elements`,
        elementDepth,
        limits.maxDepth,
      );
    }
    if (elementDepth > maxDepth) {
      maxDepth = elementDepth;
    }
    if (xml.charCodeAt(lastContent) !== 47) {
      depth = elementDepth;
    }
    cursor = close + 1;
  }

  if (depth !== 0) {
    throwSyntaxLimit();
  }

  if (budget) {
    budget.elements += nodes;
    budget.attributes += attributes;
  }
  return { elements: nodes, attributes, maxDepth };
};
