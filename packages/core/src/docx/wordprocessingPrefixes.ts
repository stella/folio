/**
 * Which prefixes a WordprocessingML part spells its namespaces with.
 *
 * A prefix is an alias: `<w:p>`, `<x:p>` and an unprefixed `<p>` under a
 * default namespace are the same paragraph when the prefix (or the default)
 * is bound to the WordprocessingML URI. The parser resolves names by URI, so
 * every such package opens. The patchers that work on part XML as a string
 * (`ensureParaIds`, the selective-save splices) find elements by their
 * literal tags instead, and a literal `<w:p` finds nothing in a part that
 * spells it `<x:p`: the patch reports success having touched nothing.
 *
 * {@link resolveWordprocessingPrefixes} is the one place those patchers ask
 * how a part spells WordprocessingML, the `w14` extensions and
 * markup compatibility (`mc`). A patcher either scans with every prefix the
 * part binds, or checks the names its literal scan needs in their namespace
 * scope. A prefix conflict on one of those names is refused.
 */

import {
  getLocalName,
  getNamespacePrefix,
  getNamespaceUri,
  NAMESPACES,
  parseXmlDocument,
  resolveAttributeNamespaceUri,
  WORDPROCESSINGML_NAMESPACE_URIS,
  type XmlElement,
} from "./xmlParser";

/** How one part spells the namespaces a string-level patcher scans for. */
export type WordprocessingPrefixes = {
  /**
   * Prefixes bound to WordprocessingML (Transitional or Strict), `""` for the
   * default namespace. Never empty: an undeclared part reads as `w`.
   */
  main: readonly string[];
  /** Prefixes bound to the `w14` extensions namespace. */
  w14: readonly string[];
  /** Whether a `w14` prefix is declared on the root (else `w14` is assumed). */
  w14Declared: boolean;
  /** Prefixes bound to markup compatibility (`mc`). */
  mc: readonly string[];
  /** Whether an `mc` prefix is declared on the root (else `mc` is assumed). */
  mcDeclared: boolean;
  /**
   * The spelling folio's serializer writes: WordprocessingML only as `w`,
   * `w14` and `mc` under their conventional prefixes. Splicing serializer
   * output into a part is only sound when this holds.
   */
  canonical: boolean;
};

export type WordprocessingPrefixResolution =
  | { type: "resolved"; prefixes: WordprocessingPrefixes }
  | { type: "unsupported"; reason: string };

const XMLNS = "xmlns";

type Declaration = { prefix: string; uri: string };

const isWhitespace = (character: string | undefined): boolean =>
  character === " " || character === "\t" || character === "\n" || character === "\r";

/** End offset (exclusive) of the tag opening at `start`, skipping quoted values. */
const tagEnd = (xml: string, start: number): number => {
  let quote: string | null = null;
  for (let index = start + 1; index < xml.length; index += 1) {
    const character = xml[index];
    if (quote !== null) {
      if (character === quote) {
        quote = null;
      }
      continue;
    }
    if (character === '"' || character === "'") {
      quote = character;
    } else if (character === ">") {
      return index + 1;
    }
  }
  return -1;
};

/** The `xmlns` / `xmlns:*` declarations written in one start tag. */
const declarationsIn = (tag: string): Declaration[] => {
  const declarations: Declaration[] = [];
  const pattern =
    /\sxmlns(?::(?<prefix>[^\s=/>]+))?\s*=\s*(?<quote>["'])(?<uri>[\s\S]*?)\k<quote>/gu;
  for (const match of tag.matchAll(pattern)) {
    declarations.push({
      prefix: match.groups?.["prefix"] ?? "",
      uri: match.groups?.["uri"] ?? "",
    });
  }
  return declarations;
};

const OPAQUE_REGIONS = [
  { open: "<!--", close: "-->" },
  { open: "<![CDATA[", close: "]]>" },
  { open: "<?", close: "?>" },
] as const;

type TagVisitor = (tag: string, isRoot: boolean) => void;

/** Visit every start tag of `xml` that declares a namespace, root first. */
const visitDeclaringStartTags = (xml: string, visit: TagVisitor): boolean => {
  let seenRoot = false;
  let pos = 0;
  scan: while (pos < xml.length) {
    const start = xml.indexOf("<", pos);
    if (start === -1) {
      break;
    }
    for (const { open, close } of OPAQUE_REGIONS) {
      if (xml.startsWith(open, start)) {
        const closeStart = xml.indexOf(close, start + open.length);
        if (closeStart === -1) {
          return false;
        }
        pos = closeStart + close.length;
        continue scan;
      }
    }
    const next = xml[start + 1];
    if (next === "!" || next === "/") {
      const end = xml.indexOf(">", start);
      if (end === -1) {
        return false;
      }
      pos = end + 1;
      continue;
    }
    const end = tagEnd(xml, start);
    if (end === -1) {
      return false;
    }
    const isRoot = !seenRoot;
    seenRoot = true;
    const tag = xml.slice(start, end);
    if (isRoot || tag.includes(XMLNS)) {
      visit(tag, isRoot);
    }
    pos = end;
  }
  return seenRoot;
};

const MAIN_URIS = WORDPROCESSINGML_NAMESPACE_URIS;
const W14_URI: string = NAMESPACES.w14;
const MC_URI: string = NAMESPACES.mc;

type NamespaceSlot = "main" | "w14" | "mc";
const NESTED_BINDINGS = { strict: "strict", scoped: "scoped" } as const;
type NestedBindings = (typeof NESTED_BINDINGS)[keyof typeof NESTED_BINDINGS];

const slotOf = (uri: string): NamespaceSlot | null => {
  if (MAIN_URIS.has(uri)) return "main";
  if (uri === W14_URI) return "w14";
  if (uri === MC_URI) return "mc";
  return null;
};

const CONVENTIONAL_PREFIX: Record<NamespaceSlot, string> = { main: "w", w14: "w14", mc: "mc" };

/**
 * Resolve the prefixes `xml` binds to WordprocessingML, `w14` and `mc`.
 *
 * Root declarations decide. A namespace the root never declares is read under
 * its conventional prefix, the tolerant reading of a malformed part. A declaration on
 * a nested element is accepted only when it repeats a binding the scan already
 * uses; anything else changes what a literal tag means partway through the
 * part and is reported as unsupported.
 */
const resolvePrefixes = (
  xml: string,
  nestedBindings: NestedBindings,
): WordprocessingPrefixResolution => {
  const bound: Record<NamespaceSlot, string[]> = { main: [], w14: [], mc: [] };
  const rootPrefixes = new Map<string, string>();
  const nested: Declaration[] = [];

  const complete = visitDeclaringStartTags(xml, (tag, isRoot) => {
    for (const declaration of declarationsIn(tag)) {
      if (!isRoot) {
        nested.push(declaration);
        continue;
      }
      rootPrefixes.set(declaration.prefix, declaration.uri);
      const slot = slotOf(declaration.uri);
      if (slot !== null) {
        bound[slot].push(declaration.prefix);
      }
    }
  });
  if (!complete) {
    return { type: "unsupported", reason: "no root element or unterminated markup" };
  }

  const declared: Record<NamespaceSlot, boolean> = {
    main: bound.main.length > 0,
    w14: bound.w14.length > 0,
    mc: bound.mc.length > 0,
  };
  for (const slot of ["main", "w14", "mc"] as const) {
    if (declared[slot]) {
      if (slot !== "main" && bound[slot].includes("")) {
        // Attributes never take the default namespace; `w14:paraId` and
        // `mc:Ignorable` need a prefix.
        return { type: "unsupported", reason: `${slot} is bound only as the default namespace` };
      }
      continue;
    }
    const conventional = CONVENTIONAL_PREFIX[slot];
    if (rootPrefixes.has(conventional)) {
      if (slot === "main") {
        return {
          type: "unsupported",
          reason: `no WordprocessingML binding; ${conventional} is bound to another namespace`,
        };
      }
      // Nothing to read under the conventional prefix, and nothing may be
      // written there either; the scan simply finds no such attributes.
      continue;
    }
    bound[slot].push(conventional);
  }

  const inUse = new Map<string, NamespaceSlot>();
  for (const slot of ["main", "w14", "mc"] as const) {
    for (const prefix of bound[slot]) {
      inUse.set(prefix, slot);
    }
  }
  for (const { prefix, uri } of nestedBindings === NESTED_BINDINGS.strict ? nested : []) {
    const slot = slotOf(uri);
    const spelling = prefix === "" ? "the default namespace" : `prefix ${prefix}`;
    if (slot !== null) {
      if (!bound[slot].includes(prefix)) {
        return {
          type: "unsupported",
          reason: `a nested element binds ${slot} under ${spelling}`,
        };
      }
      continue;
    }
    if (inUse.has(prefix)) {
      return {
        type: "unsupported",
        reason: `a nested element rebinds ${spelling} to ${uri === "" ? "no namespace" : uri}`,
      };
    }
  }

  const only = (prefixes: readonly string[], prefix: string): boolean =>
    prefixes.length === 1 && prefixes[0] === prefix;
  // An empty list means the conventional prefix is bound to another
  // namespace, so the serializer's `w14:` / `mc:` would not mean what it says.
  const canonical = only(bound.main, "w") && only(bound.w14, "w14") && only(bound.mc, "mc");

  return {
    type: "resolved",
    prefixes: {
      main: bound.main,
      w14: bound.w14,
      w14Declared: declared.w14,
      mc: bound.mc,
      mcDeclared: declared.mc,
      canonical,
    },
  };
};

export const resolveWordprocessingPrefixes = (xml: string): WordprocessingPrefixResolution =>
  resolvePrefixes(xml, NESTED_BINDINGS.strict);

/** Root spelling for a selective scan whose relevant nested names are checked separately. */
export const resolveSelectiveScanPrefixes = (xml: string): WordprocessingPrefixResolution =>
  resolvePrefixes(xml, NESTED_BINDINGS.scoped);

const escapeRegExp = (value: string): string => value.replace(/[.*+?^${}()|[\]\\]/gu, "\\$&");

/**
 * Whether a patcher that reads `<w:…>` literally sees every element it looks
 * for in `xml`, and may splice the serializer's `w:` markup into it.
 *
 * True for the canonical spelling. Also true when the part binds
 * WordprocessingML under an extra prefix besides `w` but never spells the
 * names the patcher scans under that alias, or binds markup compatibility
 * under another prefix. Relevant nested names are checked in their own scope.
 */
export const splicesAsCanonical = (xml: string, localNames: readonly string[]): boolean => {
  const resolution = resolveWordprocessingPrefixes(xml);
  if (resolution.type === "resolved" && resolution.prefixes.canonical) return true;
  const relaxed = resolveSelectiveScanPrefixes(xml);
  if (relaxed.type !== "resolved") return false;
  const { prefixes } = relaxed;
  const only = (list: readonly string[], prefix: string): boolean =>
    list.length === 1 && list[0] === prefix;
  const aliases = prefixes.main.filter((prefix) => prefix !== "w");
  if (
    !prefixes.main.includes("w") ||
    aliases.includes("") ||
    !only(prefixes.w14, "w14") ||
    prefixes.mc.length === 0
  ) {
    return false;
  }
  const names = localNames.map(escapeRegExp).join("|");
  if (
    !aliases.every(
      (alias) => !new RegExp(`[<\\s/]${escapeRegExp(alias)}:(?:${names})[\\s/>=]`, "u").test(xml),
    )
  ) {
    return false;
  }

  // A nested rebind matters only when one of the literal names the splice
  // scans occurs under it. Resolve those names in their actual XML scope.
  const root = parseXmlDocument(xml);
  if (!root) return false;
  const mainNames = new Set(localNames.filter((name) => name !== "paraId" && name !== "textId"));
  const idNames = new Set(["paraId", "textId"]);
  const mcNames = new Set(["Fallback", "AlternateContent"]);
  const visit = (element: XmlElement): boolean => {
    if (element.type !== "element") return true;
    const name = element.name ?? "";
    const local = getLocalName(name);
    const prefix = getNamespacePrefix(name) ?? "";
    const uri = getNamespaceUri(element) ?? "";
    if (
      (prefix === "mc" && uri !== MC_URI) ||
      (mainNames.has(local) && WORDPROCESSINGML_NAMESPACE_URIS.has(uri) !== (prefix === "w")) ||
      (mcNames.has(local) && (uri === MC_URI) !== prefixes.mc.includes(prefix))
    ) {
      return false;
    }
    for (const attribute of Object.keys(element.attributes ?? {})) {
      const attributePrefix = getNamespacePrefix(attribute);
      if (attributePrefix === "mc" && resolveAttributeNamespaceUri(element, attribute) !== MC_URI) {
        return false;
      }
      if (!idNames.has(getLocalName(attribute))) continue;
      if (
        (resolveAttributeNamespaceUri(element, attribute) === W14_URI) !==
        (attributePrefix === "w14")
      ) {
        return false;
      }
    }
    return (element.elements ?? []).every(visit);
  };
  return visit(root);
};

/** The names the paragraph splices scan for: paragraphs, their containers and ids. */
export const PARAGRAPH_SCAN_NAMES = ["p", "tc", "txbxContent", "paraId", "textId"] as const;

/** Whether `xml` spells WordprocessingML the way folio's serializer does. */
export const hasCanonicalWordprocessingPrefixes = (xml: string): boolean => {
  const resolution = resolveWordprocessingPrefixes(xml);
  return resolution.type === "resolved" && resolution.prefixes.canonical;
};

/** `<p`, `<w:p`, … — the open-tag literal of `localName` under `prefix`. */
export const openTagLiteral = (prefix: string, localName: string): string =>
  prefix === "" ? `<${localName}` : `<${prefix}:${localName}`;

/** `</p>`, `</w:p>`, … — the close-tag literal of `localName` under `prefix`. */
export const closeTagLiteral = (prefix: string, localName: string): string =>
  prefix === "" ? `</${localName}>` : `</${prefix}:${localName}>`;

/**
 * Whether `xml` opens `localName` under one of `prefixes` at `start`, and the
 * length of the literal that matched (`-1` when none did). The character after
 * the name must end it, so `<w:pPr` is not `<w:p`.
 */
export const matchOpenTag = (
  xml: string,
  start: number,
  prefixes: readonly string[],
  localName: string,
): number => {
  for (const prefix of prefixes) {
    const literal = openTagLiteral(prefix, localName);
    if (!xml.startsWith(literal, start)) continue;
    const after = xml[start + literal.length];
    if (after === ">" || after === "/" || isWhitespace(after)) {
      return literal.length;
    }
  }
  return -1;
};

/** The length of the close tag of `localName` under one of `prefixes` at `start`, or -1. */
export const matchCloseTag = (
  xml: string,
  start: number,
  prefixes: readonly string[],
  localName: string,
): number => {
  for (const prefix of prefixes) {
    const literal = closeTagLiteral(prefix, localName);
    if (xml.startsWith(literal, start)) {
      return literal.length;
    }
  }
  return -1;
};
