/**
 * Namespace declarations for parts folio rebuilds from the model.
 *
 * A rebuilt part carries content the parser preserved verbatim: raw property
 * XML, text-box bodies, drawing extensions, unmodeled attributes. That content
 * can use any prefix its producer bound on the source part's root, so a
 * hand-maintained declaration list on the serializer silently drops bindings
 * and emits an unbound prefix, which makes the part malformed and the package
 * unopenable. The declarations are therefore derived: every prefix the
 * assembled part actually uses is resolved through {@link OOXML_NAMESPACES},
 * falling back to what the source part bound, and a prefix that resolves to
 * nothing fails the save instead of being written unbound.
 */

import { OOXML_NS, type OoxmlPrefix } from "@stll/docx-utils";
import { TaggedError } from "better-result";

import { escapeXml } from "./xmlUtils";

/**
 * A namespace a rebuilt WordprocessingML part may declare.
 *
 * `ignorable` marks an extension that decorates baseline elements in place
 * (`w14:paraId`, `w16du:dateUtc`, `wp14:sizeRelH`): a consumer that predates it
 * has to be told through `mc:Ignorable` to skip it, or it rejects the part.
 * Extensions reachable only inside an `mc:AlternateContent` choice or an
 * `a:graphicData` payload (`wps`, `wpg`, `wne`, `cx*`) are already gated by
 * that wrapper, so they stay out of the attribute, as mainstream producers
 * leave them.
 */
export type OoxmlNamespace = {
  readonly uri: string;
  readonly ignorable: boolean;
};

/** Packaging prefixes; they never appear inside a WordprocessingML part. */
type PackagingPrefix = "ct" | "pr";

/**
 * A prefix folio can bind on a rebuilt part without consulting the source.
 *
 * Every prefix the parser resolves is a member, so a prefix the parser
 * understands can always be written back, and {@link OOXML_NAMESPACES} is total
 * over the union: a member with no entry, or an entry with no member, does not
 * compile.
 */
export type OoxmlNamespacePrefix =
  | Exclude<OoxmlPrefix, PackagingPrefix>
  | "cx"
  | "cx1"
  | "cx2"
  | "cx3"
  | "cx4"
  | "cx5"
  | "cx6"
  | "cx7"
  | "cx8"
  | "aink"
  | "am3d"
  | "oel"
  | "w10"
  | "w16"
  | "w16cex"
  | "w16cid"
  | "w16du"
  | "w16sdtdh"
  | "w16se"
  | "wpi"
  | "wne";

/**
 * Every namespace a rebuilt part can declare, in the order producers write them.
 *
 * URIs shared with the parser are taken from its table rather than retyped, so
 * the two cannot drift.
 */
export const OOXML_NAMESPACES: Readonly<Record<OoxmlNamespacePrefix, OoxmlNamespace>> = {
  wpc: { uri: OOXML_NS.wpc, ignorable: false },
  cx: { uri: "http://schemas.microsoft.com/office/drawing/2014/chartex", ignorable: false },
  cx1: { uri: "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex", ignorable: false },
  cx2: { uri: "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex", ignorable: false },
  cx3: { uri: "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex", ignorable: false },
  cx4: { uri: "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex", ignorable: false },
  cx5: { uri: "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex", ignorable: false },
  cx6: { uri: "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex", ignorable: false },
  cx7: { uri: "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex", ignorable: false },
  cx8: { uri: "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex", ignorable: false },
  mc: { uri: OOXML_NS.mc, ignorable: false },
  aink: { uri: "http://schemas.microsoft.com/office/drawing/2016/ink", ignorable: false },
  am3d: { uri: "http://schemas.microsoft.com/office/drawing/2017/model3d", ignorable: false },
  o: { uri: OOXML_NS.o, ignorable: false },
  oel: { uri: "http://schemas.microsoft.com/office/2019/extlst", ignorable: false },
  r: { uri: OOXML_NS.r, ignorable: false },
  m: { uri: OOXML_NS.m, ignorable: false },
  v: { uri: OOXML_NS.v, ignorable: false },
  wp14: { uri: OOXML_NS.wp14, ignorable: true },
  wp: { uri: OOXML_NS.wp, ignorable: false },
  w10: { uri: "urn:schemas-microsoft-com:office:word", ignorable: false },
  w: { uri: OOXML_NS.w, ignorable: false },
  w14: { uri: OOXML_NS.w14, ignorable: true },
  w15: { uri: OOXML_NS.w15, ignorable: true },
  w16cex: { uri: "http://schemas.microsoft.com/office/word/2018/wordml/cex", ignorable: true },
  w16cid: { uri: "http://schemas.microsoft.com/office/word/2016/wordml/cid", ignorable: true },
  w16: { uri: "http://schemas.microsoft.com/office/word/2018/wordml", ignorable: true },
  w16du: { uri: "http://schemas.microsoft.com/office/word/2023/wordml/word16du", ignorable: true },
  w16sdtdh: {
    uri: "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash",
    ignorable: true,
  },
  w16se: { uri: "http://schemas.microsoft.com/office/word/2015/wordml/symex", ignorable: true },
  wpg: { uri: OOXML_NS.wpg, ignorable: false },
  wpi: { uri: "http://schemas.microsoft.com/office/word/2010/wordprocessingInk", ignorable: false },
  wne: { uri: "http://schemas.microsoft.com/office/word/2006/wordml", ignorable: false },
  wps: { uri: OOXML_NS.wps, ignorable: false },
  a: { uri: OOXML_NS.a, ignorable: false },
  pic: { uri: OOXML_NS.pic, ignorable: false },
};

/** {@link OOXML_NAMESPACES} keyed for lookup by a prefix read out of markup. */
const NAMESPACE_TABLE: ReadonlyMap<string, OoxmlNamespace> = new Map(
  Object.entries(OOXML_NAMESPACES),
);

/** A rebuilt part uses a prefix that resolves to no namespace URI. */
export class UnboundNamespacePrefixError extends TaggedError("UnboundNamespacePrefixError")<{
  message: string;
  partPath: string;
  prefixes: readonly string[];
}> {}

const XMLNS_ATTRIBUTE_PREFIX = "xmlns:";
/** Bound by the XML specification itself; never declared. */
const RESERVED_PREFIX = "xml";

const isWhitespace = (character: string): boolean =>
  character === " " || character === "\t" || character === "\n" || character === "\r";

const isNameBoundary = (character: string): boolean =>
  isWhitespace(character) || character === ">" || character === "/" || character === "=";

const namePrefix = (qualifiedName: string): string | undefined => {
  const colon = qualifiedName.indexOf(":");
  if (colon <= 0) {
    return undefined;
  }
  const prefix = qualifiedName.slice(0, colon);
  return prefix === RESERVED_PREFIX ? undefined : prefix;
};

type ScannedNames = {
  /** Prefixes carried by element and attribute names. */
  used: ReadonlySet<string>;
  /** `xmlns:*` bindings found in the scanned markup, latest occurrence winning. */
  declared: ReadonlyMap<string, string>;
};

/**
 * Collect prefixes from element and attribute names in one quote-aware pass.
 *
 * Attribute values are skipped rather than matched, so a URI or a text value
 * containing a colon cannot be mistaken for a namespace prefix.
 */
const scanXmlNames = (xml: string, rootTagOnly: boolean): ScannedNames => {
  const used = new Set<string>();
  const declared = new Map<string, string>();
  const length = xml.length;
  let index = 0;

  while (index < length) {
    const tagStart = xml.indexOf("<", index);
    if (tagStart === -1) {
      break;
    }
    index = tagStart + 1;

    if (xml.startsWith("?", index)) {
      const end = xml.indexOf("?>", index);
      if (end === -1) {
        break;
      }
      index = end + 2;
      continue;
    }
    if (xml.startsWith("!--", index)) {
      const end = xml.indexOf("-->", index);
      if (end === -1) {
        break;
      }
      index = end + 3;
      continue;
    }
    if (xml.startsWith("![CDATA[", index)) {
      const end = xml.indexOf("]]>", index);
      if (end === -1) {
        break;
      }
      index = end + 3;
      continue;
    }
    if (xml.startsWith("!", index)) {
      const end = xml.indexOf(">", index);
      if (end === -1) {
        break;
      }
      index = end + 1;
      continue;
    }
    if (xml.startsWith("/", index)) {
      index += 1;
    }

    const nameStart = index;
    while (index < length && !isNameBoundary(xml[index] ?? ">")) {
      index += 1;
    }
    const elementPrefix = namePrefix(xml.slice(nameStart, index));
    if (elementPrefix) {
      used.add(elementPrefix);
    }

    while (index < length) {
      const character = xml[index] ?? ">";
      if (isWhitespace(character) || character === "/") {
        index += 1;
        continue;
      }
      if (character === ">") {
        index += 1;
        break;
      }

      const attributeStart = index;
      while (index < length && !isNameBoundary(xml[index] ?? ">")) {
        index += 1;
      }
      const attributeName = xml.slice(attributeStart, index);

      while (index < length && isWhitespace(xml[index] ?? ">")) {
        index += 1;
      }
      let value = "";
      if (xml[index] === "=") {
        index += 1;
        while (index < length && isWhitespace(xml[index] ?? ">")) {
          index += 1;
        }
        const quote = xml[index];
        if (quote === '"' || quote === "'") {
          const valueEnd = xml.indexOf(quote, index + 1);
          if (valueEnd === -1) {
            index = length;
            break;
          }
          value = xml.slice(index + 1, valueEnd);
          index = valueEnd + 1;
        }
      }

      if (attributeName.startsWith(XMLNS_ATTRIBUTE_PREFIX)) {
        declared.set(attributeName.slice(XMLNS_ATTRIBUTE_PREFIX.length), value);
        continue;
      }
      const attributePrefix = namePrefix(attributeName);
      if (attributePrefix) {
        used.add(attributePrefix);
      }
    }

    if (rootTagOnly) {
      break;
    }
  }

  return { used, declared };
};

/**
 * The `xmlns:*` bindings a source part declared on its root element.
 *
 * Reused when the part is rebuilt so a producer-specific prefix that
 * {@link OOXML_NAMESPACES} does not know keeps the URI the document gave it.
 */
export const readRootNamespaceBindings = (partXml: string): ReadonlyMap<string, string> =>
  scanXmlNames(partXml, true).declared;

/** The part a rebuild replaces: where it lives, and what its root element bound. */
export type SourcePart = {
  path: string;
  bindings: ReadonlyMap<string, string>;
};

export type PartElementOptions = {
  /** Package-relative path of the part, reported when a prefix is unbound. */
  partPath: string;
  /** Qualified name of the root element, for example `w:document`. */
  rootName: string;
  /** Root-element attributes other than the namespace declarations. */
  rootAttributes?: string;
  /**
   * Prefixes declared whether or not the body uses them, so the part keeps a
   * stable shape across saves and unmodeled extensions have somewhere to land.
   */
  baselinePrefixes: readonly OoxmlNamespacePrefix[];
  /** Root bindings of the part being replaced, when the save has the source. */
  sourceBindings: ReadonlyMap<string, string> | undefined;
  /** Assembled part content, everything between the root tags. */
  body: string;
};

/**
 * Wrap an assembled part body in a root element that declares every prefix the
 * part uses and lists the ignorable ones in `mc:Ignorable`.
 */
export const serializePartElement = ({
  partPath,
  rootName,
  rootAttributes,
  baselinePrefixes,
  sourceBindings,
  body,
}: PartElementOptions): string => {
  const openingTag = `<${rootName}${rootAttributes ? ` ${rootAttributes}` : ""}`;
  const fromRoot = scanXmlNames(`${openingTag}/>`, true);
  const { used, declared } = scanXmlNames(body, false);
  const required = new Set<string>([...fromRoot.used, ...used]);

  const bindings = new Map<string, string>();
  for (const prefix of baselinePrefixes) {
    bindings.set(prefix, OOXML_NAMESPACES[prefix].uri);
  }

  const unbound: string[] = [];
  for (const prefix of required) {
    if (bindings.has(prefix)) {
      continue;
    }
    const uri = NAMESPACE_TABLE.get(prefix)?.uri ?? sourceBindings?.get(prefix);
    if (uri !== undefined) {
      bindings.set(prefix, uri);
      continue;
    }
    // An inner element may bind the prefix for its own subtree; only a prefix
    // nothing binds anywhere would be written unbound.
    if (!declared.has(prefix)) {
      unbound.push(prefix);
    }
  }

  if (unbound.length > 0) {
    const prefixes = unbound.sort();
    throw new UnboundNamespacePrefixError({
      message: `${partPath} uses namespace prefixes bound to no URI: ${prefixes.join(", ")}`,
      partPath,
      prefixes,
    });
  }

  const hasIgnorable = [...bindings.keys()].some(
    (prefix) => NAMESPACE_TABLE.get(prefix)?.ignorable === true,
  );
  if (hasIgnorable) {
    // The attribute itself needs the markup-compatibility binding.
    bindings.set("mc", OOXML_NAMESPACES.mc.uri);
  }

  // Baseline order first, then the table's order for what the body added, then
  // source-only prefixes: a part whose body needs nothing extra keeps the exact
  // declaration list it had before, so an unrelated edit does not rewrite it.
  const baseline = new Set<string>(baselinePrefixes);
  const extra = new Set([...bindings.keys()].filter((prefix) => !baseline.has(prefix)));
  const orderedPrefixes = [
    ...baselinePrefixes,
    ...[...NAMESPACE_TABLE.keys()].filter((prefix) => extra.has(prefix)),
    ...[...extra].filter((prefix) => !NAMESPACE_TABLE.has(prefix)).sort(),
  ];
  const attributes = orderedPrefixes.map(
    // SAFETY: every ordered prefix came from `bindings`.
    (prefix) => `xmlns:${prefix}="${escapeXml(bindings.get(prefix)!)}"`,
  );
  const ignorable = orderedPrefixes.filter(
    (prefix) => NAMESPACE_TABLE.get(prefix)?.ignorable === true,
  );
  if (ignorable.length > 0) {
    attributes.push(`mc:Ignorable="${ignorable.join(" ")}"`);
  }

  return `${openingTag} ${attributes.join(" ")}>${body}</${rootName}>`;
};
