/**
 * The same package, spelled with different namespace prefixes.
 *
 * A prefix is an alias for its namespace URI, so `<w:p>`, `<x:p>` and `<p>`
 * under a WordprocessingML default namespace are one element. Every fixture
 * folio tests with uses the conventional prefixes, which is how a patcher
 * that reads XML as a string can match `<w:p` literally and still pass. These
 * helpers rewrite any fixture into equivalent spellings, and compare parts by
 * what they mean rather than how they are spelled, so a test can hold a
 * string-level patcher to "same result or an explicit refusal".
 */

import JSZip from "jszip";

import { canonicalJson } from "../../utils/canonicalJson";
import {
  NAMESPACES,
  getLocalName,
  getNamespaceUri,
  parseXmlDocument,
  resolveAttributeNamespaceUri,
  type XmlElement,
} from "../xmlParser";

/**
 * - `alias`: WordprocessingML under `x:`, the `w14` extensions under `x14:`.
 * - `default`: WordprocessingML elements unprefixed under a default namespace;
 *   its attributes (which never take a default namespace) under `x:`, and the
 *   `w14` extensions under `x14:`.
 * - `w14-alias`: only the `w14` extensions move, to `x14:`; `w:` stays.
 */
export type PrefixVariant = "alias" | "default" | "w14-alias";

export const PREFIX_VARIANTS: readonly PrefixVariant[] = ["alias", "default", "w14-alias"];

const W_URI: string = NAMESPACES.w;
const W14_URI: string = NAMESPACES.w14;

/** Attributes whose value is a list of prefixes (or prefixed names) rather than data. */
const PREFIX_LIST_ATTRIBUTES = new Set([
  "Ignorable",
  "ProcessContent",
  "PreserveElements",
  "PreserveAttributes",
  "MustUnderstand",
  "Requires",
]);

const XML_TOKEN =
  /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>|<!DOCTYPE[^>]*>|<(?:"[^"]*"|'[^']*'|[^'">])*>/gu;
const ATTRIBUTE = /(?<space>\s+)(?<name>[^\s=/>]+)(?<eq>\s*=\s*)(?<value>"[^"]*"|'[^']*')/gu;

type Renames = {
  /** Prefix used for WordprocessingML elements (`""` for the default namespace). */
  elementPrefix: string;
  /** Prefix used for WordprocessingML attributes. */
  attributePrefix: string;
  w14Prefix: string;
};

const RENAMES: Record<PrefixVariant, Renames> = {
  alias: { elementPrefix: "x", attributePrefix: "x", w14Prefix: "x14" },
  default: { elementPrefix: "", attributePrefix: "x", w14Prefix: "x14" },
  "w14-alias": { elementPrefix: "w", attributePrefix: "w", w14Prefix: "x14" },
};

const qualify = (prefix: string, local: string): string =>
  prefix === "" ? local : `${prefix}:${local}`;

const renamePrefixToken = (token: string, renames: Renames): string => {
  const colon = token.indexOf(":");
  const prefix = colon === -1 ? token : token.slice(0, colon);
  const rest = colon === -1 ? "" : token.slice(colon);
  return renamedPrefix(prefix, renames.attributePrefix, renames) + rest;
};

/** `w` → the variant's WordprocessingML prefix, `w14` → its `w14` prefix, others kept. */
const renamedPrefix = (prefix: string, wordPrefix: string, renames: Renames): string => {
  if (prefix === "w") return wordPrefix;
  if (prefix === "w14") return renames.w14Prefix;
  return prefix;
};

const rewriteTag = (tag: string, renames: Renames): string => {
  const closing = tag.startsWith("</");
  const nameStart = closing ? 2 : 1;
  let nameEnd = nameStart;
  while (nameEnd < tag.length && !/[\s/>]/u.test(tag[nameEnd] ?? ">")) {
    nameEnd += 1;
  }
  const name = tag.slice(nameStart, nameEnd);
  const colon = name.indexOf(":");
  const renamedName =
    colon === -1
      ? name
      : qualify(
          renamedPrefix(name.slice(0, colon), renames.elementPrefix, renames),
          name.slice(colon + 1),
        );
  if (closing) {
    return `</${renamedName}${tag.slice(nameEnd)}`;
  }
  const attributes = tag.slice(nameEnd).replace(ATTRIBUTE, (...args: unknown[]) => {
    const groups = args.at(-1) as Record<string, string>;
    const {
      space,
      name: attributeName,
      eq,
      value,
    } = groups as {
      space: string;
      name: string;
      eq: string;
      value: string;
    };
    if (attributeName === "xmlns:w") {
      return renames.elementPrefix === ""
        ? `${space}xmlns${eq}${value} xmlns:${renames.attributePrefix}${eq}${value}`
        : `${space}xmlns:${renames.elementPrefix}${eq}${value}`;
    }
    if (attributeName === "xmlns:w14") {
      return `${space}xmlns:${renames.w14Prefix}${eq}${value}`;
    }
    const renamedAttribute = renamePrefixToken(attributeName, renames);
    const local = attributeName.slice(attributeName.indexOf(":") + 1);
    if (PREFIX_LIST_ATTRIBUTES.has(local)) {
      const quote = value[0] ?? '"';
      const tokens = value
        .slice(1, -1)
        .split(/(\s+)/u)
        .map((token) => (/^\s*$/u.test(token) ? token : renamePrefixToken(token, renames)));
      return `${space}${renamedAttribute}${eq}${quote}${tokens.join("")}${quote}`;
    }
    return `${space}${renamedAttribute}${eq}${value}`;
  });
  return `<${renamedName}${attributes}`;
};

/**
 * `xml` respelled as `variant`, or null when the part cannot be respelled
 * without changing what it means: it does not bind `w` to WordprocessingML
 * on its root, already uses one of the target prefixes, already declares a
 * default namespace, or (for `default`) has an unprefixed element that would
 * fall into WordprocessingML.
 */
export const rewritePartPrefixes = (xml: string, variant: PrefixVariant): string | null => {
  const renames = RENAMES[variant];
  const root = parseXmlDocument(xml);
  if (!root?.name?.startsWith("w:") || getNamespaceUri(root) !== W_URI) {
    return null;
  }
  if (/\sxmlns\s*=/u.test(xml) || /[<\s/](?:x|x14):/u.test(xml)) {
    return null;
  }
  const w14Bound = /\sxmlns:w14\s*=\s*["']([^"']*)["']/u.exec(xml)?.[1];
  if (w14Bound !== undefined && w14Bound !== W14_URI) {
    return null;
  }
  let unprefixedElement = false;
  const rewritten = xml.replace(XML_TOKEN, (tag) => {
    if (tag.startsWith("<!") || tag.startsWith("<?")) {
      return tag;
    }
    if (!/^<\/?[^\s/>]*:/u.test(tag)) {
      unprefixedElement = true;
    }
    return rewriteTag(tag, renames);
  });
  if (variant === "default" && unprefixedElement) {
    return null;
  }
  return rewritten;
};

/** Parts a variant respells: every WordprocessingML XML part under `word/`. */
const isWordXmlPart = (path: string): boolean => {
  const lower = path.toLowerCase();
  return lower.startsWith("word/") && lower.endsWith(".xml") && !lower.includes("/_rels/");
};

/**
 * `docx` with every WordprocessingML part respelled as `variant`, and the
 * paths it respelled. Null when not even `word/document.xml` can be.
 */
export const rewritePackagePrefixes = async (
  docx: Uint8Array | ArrayBuffer,
  variant: PrefixVariant,
): Promise<{ docx: Uint8Array; rewritten: string[] } | null> => {
  const zip = await JSZip.loadAsync(docx);
  const rewritten: string[] = [];
  for (const [path, file] of Object.entries(zip.files)) {
    if (file.dir || !isWordXmlPart(path)) continue;
    // oxlint-disable-next-line no-await-in-loop -- a handful of parts per fixture
    const xml = await file.async("text");
    const respelled = rewritePartPrefixes(xml, variant);
    if (respelled === null) continue;
    zip.file(path, respelled);
    rewritten.push(path);
  }
  if (!rewritten.some((path) => path.toLowerCase() === "word/document.xml")) {
    return null;
  }
  return { docx: await zip.generateAsync({ type: "uint8array" }), rewritten };
};

type CanonicalNode =
  | { text: string }
  | { name: string; attributes: [string, string][]; children: CanonicalNode[] };

const IDENTITY_ATTRIBUTES = new Set(["paraId", "textId"]);
const TEXT_ELEMENTS = new Set(["t", "delText", "instrText", "delInstrText"]);

/**
 * A part as namespace URIs and local names, prefixes and declarations gone,
 * with paragraph identity values replaced by the order they first appear in:
 * two runs that minted different ids for the same paragraphs compare equal.
 */
export const canonicalPartTree = (xml: string): string => {
  const root = parseXmlDocument(xml);
  if (!root) {
    throw new Error("canonicalPartTree: part does not parse");
  }
  const identities = new Map<string, string>();
  const identity = (value: string): string => {
    const key = value.toUpperCase();
    let ordinal = identities.get(key);
    if (ordinal === undefined) {
      ordinal = `#${identities.size + 1}`;
      identities.set(key, ordinal);
    }
    return ordinal;
  };
  const visit = (element: XmlElement): CanonicalNode => {
    if (element.type === "text" || element.type === "cdata") {
      return { text: String(element.text ?? element.cdata ?? "") };
    }
    const attributes: [string, string][] = [];
    for (const [name, raw] of Object.entries(element.attributes ?? {})) {
      if (raw === undefined || name === "xmlns" || name.startsWith("xmlns:")) continue;
      const local = name.slice(name.indexOf(":") + 1);
      const uri = resolveAttributeNamespaceUri(element, name) ?? "";
      let value = String(raw);
      if (IDENTITY_ATTRIBUTES.has(local)) {
        value = identity(value);
      } else if (PREFIX_LIST_ATTRIBUTES.has(local)) {
        value = value
          .split(/\s+/u)
          .filter((token) => token.length > 0)
          .map((token) => {
            const [prefix = "", rest = ""] = token.split(":");
            const resolved = resolveAttributeNamespaceUri(element, `${prefix}:_`) ?? prefix;
            return rest === "" ? resolved : `{${resolved}}${rest}`;
          })
          .sort()
          .join(" ");
      }
      attributes.push([`{${uri}}${local}`, value]);
    }
    attributes.sort(([a], [b]) => a.localeCompare(b));
    const localName = getLocalName(element.name ?? "");
    // Whitespace between elements is layout, not content, except inside the
    // elements that hold text.
    const holdsText = TEXT_ELEMENTS.has(localName);
    return {
      name: `{${getNamespaceUri(element) ?? ""}}${localName}`,
      attributes,
      children: (element.elements ?? [])
        .filter(
          (child) =>
            child.type !== "comment" &&
            child.type !== "instruction" &&
            (holdsText || child.type !== "text" || String(child.text ?? "").trim() !== ""),
        )
        .map(visit),
    };
  };
  return canonicalJson(visit(root));
};
