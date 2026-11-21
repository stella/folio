const EnsureParaIdsErrorBase: TaggedErrorClass<"EnsureParaIdsError"> =
  TaggedError("EnsureParaIdsError");

/**
 * Headless `w14:paraId` normalization for a `.docx` buffer.
 *
 * Word 2010+ stamps every paragraph with a `w14:paraId`, but Google Docs
 * exports, LibreOffice, python-docx, and docx4j generally do not. Paragraphs
 * without one are invisible to everything that anchors on paraId (block ids,
 * comment threads, AI-edit targeting) and fall back to positional `seq-NNNN`
 * ids that renumber after structural edits. Hosts call {@link ensureParaIds}
 * once at ingest so every stored version has full id coverage before any
 * snapshot or external anchor is created.
 *
 * IDs remain stable in folio and identity-preserving DOCX round-trips.
 * Microsoft Word can establish a new `w14:docId` and replace all paragraph
 * IDs on its first save of a non-Word-produced package; callers that ingest
 * such an externally edited version must normalize it again and cannot assume
 * the pre-save IDs still match.
 *
 * The pass patches part XML in place (string splices, no model round-trip),
 * so documents carrying features folio's parser does not model come back
 * with those features byte-identical. Contract:
 *
 * - Parts covered: the officeDocument relationship target, the header,
 *   footer, footnotes, and endnotes parts its relationships name (wherever
 *   they live), and `word/header*.xml`, `word/footer*.xml`,
 *   `word/footnotes.xml`, `word/endnotes.xml`. Paragraphs nested in table
 *   cells and in `mc:Choice` text boxes are plain `<w:p>` elements inside
 *   those parts and are covered by the same scan. The comments part mints its
 *   own deterministic paraIds at save time (see `commentSerializer`) and is
 *   left untouched; its ids still count toward uniqueness.
 * - Paragraphs inside `mc:Fallback` are never modified: Word regenerates the
 *   fallback branch (duplicating the `mc:Choice` ids) on save, so stamping or
 *   deduplicating there would churn on every Word round-trip.
 * - Existing ids are preserved. Only these cases get a fresh id: a missing
 *   `paraId`, the reserved all-zero value (Word reads `00000000` as "no id"),
 *   a value that is not eight hex digits (the parser drops it as no id),
 *   and a duplicate of an id already seen earlier in the scan (first
 *   occurrence keeps it — the same rule `ParaIdAllocatorExtension` applies in
 *   the editor). `w14:textId` is written alongside a newly minted `paraId`
 *   (same value) and never touched otherwise; it is a text-revision marker,
 *   not identity.
 * - Fresh ids are deterministic (`deterministicHexId` over the part path,
 *   the paragraph's own markup, and how many earlier paragraphs of the part
 *   repeat that markup), so the pass is a pure function of the input bytes:
 *   retrying an ingest produces identical output. A paragraph an edit
 *   elsewhere left as it was keeps its id in the edited version.
 * - Each patched part's root element gets `xmlns:w14` / `xmlns:mc`
 *   declarations and a `mc:Ignorable` listing `w14` when missing — non-Word
 *   producers declare neither, and absent `mc:Ignorable` handling is what
 *   makes pre-2010 consumers choke on the new attributes.
 * - Prefixes are aliases: elements and attributes resolve against their
 *   in-scope namespace URI, including Strict and Transitional profiles and
 *   nested rebinding. Minted attributes use a root prefix with no conflicting
 *   nested binding; a fresh prefix is declared when necessary.
 * - A DrawingML graphic with a missing namespace declaration is repaired in
 *   its own scope when its drawing container establishes the OOXML profile.
 * - Idempotent: a document that already has full coverage is returned as the
 *   original bytes, untouched (`alreadyComplete: true`). A body with
 *   paragraphs the scan did not see is refused, never reported complete.
 * - Digitally signed packages are returned untouched when already complete.
 *   When normalization would rewrite the package, it fails unless the caller
 *   explicitly allows signature invalidation after warning the user.
 */
import { Result, TaggedError, type TaggedErrorClass } from "better-result";

import {
  getXmlSourceRange,
  malformedXmlReferenceError,
  parseStreamingXmlWithSourceRanges,
  repairMissingDrawingNamespace,
  scanStreamingXmlElements,
} from "@stll/docx-core/xml";
import { resolveNamespaceUri } from "@stll/docx-core/xml/namespaces";
import { ZipArchive } from "@stll/docx-core/zip";

import { deterministicHexId, isValidHexId } from "../utils/hexId";
import { resolvePackageRelationshipTarget } from "./packageParts";
import { spliceXml } from "./selectiveXmlPatch";
import { loadDocxArchive } from "./server/boundedArchive";
import {
  getAttribute,
  getChildElements,
  getLocalName,
  getNamespaceUri,
  OFFICE_RELATIONSHIP_NAMESPACE_URIS,
  parseXmlDocument,
  resolveAttributeNamespaceUri,
  WORDPROCESSINGML_NAMESPACE_URIS,
  type XmlElement,
} from "./xmlParser";

export const ENSURE_PARA_IDS_REASONS = {
  NORMALIZATION_FAILED: "normalization-failed",
  NAMESPACE_INVALID: "namespace-invalid",
  SIGNED_PACKAGE: "signed-package",
} as const;

/** A malformed or unsupported package prevented paragraph-ID normalization. */
export class EnsureParaIdsError extends EnsureParaIdsErrorBase<{
  message: string;
  reason: (typeof ENSURE_PARA_IDS_REASONS)[keyof typeof ENSURE_PARA_IDS_REASONS];
  cause?: unknown;
}> {}

/** Counts and normalized bytes returned by {@link ensureParaIds}. */
export type EnsureParaIdsResult = {
  /** The normalized `.docx`; the input bytes verbatim when `alreadyComplete`. */
  docx: Uint8Array;
  /** Paragraphs that received a paraId (missing or all-zero before). */
  assigned: number;
  /** Duplicate paraIds reassigned (the first occurrence keeps the id). */
  deduplicated: number;
  /** True when the input already had full, unique coverage. */
  alreadyComplete: boolean;
  /**
   * Every id this pass wrote (assigned or deduplicated), in scan order. An id
   * absent from this list was already the package's own.
   */
  mintedParaIds: readonly string[];
};

/** Controls mutation of package metadata that has security implications. */
export type EnsureParaIdsOptions = {
  /**
   * Allow normalization to invalidate existing OPC digital signatures.
   * Callers must warn the user before opting in.
   */
  allowSignedPackageMutation?: boolean;
};

const W14_NAMESPACE_URI =
  "http://schemas.microsoft.com/office/word/2010/wordml";
const MC_NAMESPACE_URI =
  "http://schemas.openxmlformats.org/markup-compatibility/2006";
const MC_IGNORABLE_W14 = "w14";
const DIGITAL_SIGNATURE_PART_PREFIX = "_xmlsignatures/";
/** Word reads an all-zero `w14:paraId` as "no id assigned". */
const RESERVED_ZERO_ID_PATTERN = /^0*$/u;

const PACKAGE_RELATIONSHIP_URI =
  "http://schemas.openxmlformats.org/package/2006/relationships";
/** Relationship types, under each office namespace, whose targets are editable stories. */
const STORY_RELATIONSHIP_TYPES = [
  "header",
  "footer",
  "footnotes",
  "endnotes",
] as const;
const TARGET_PART_PATTERNS = [
  /^word\/header\d*\.xml$/u,
  /^word\/footer\d*\.xml$/u,
  /^word\/footnotes\.xml$/u,
  /^word\/endnotes\.xml$/u,
];

/** Reserve package ids even in parts whose content the pass leaves untouched. */
const ANY_PARA_ID_PATTERN =
  /\s[^\s=<>/"':]+:paraId\s*=\s*(?<quote>["'])(?<id>[\s\S]*?)\k<quote>/gu;

type SpliceEdit = { start: number; end: number; text: string };

/**
 * Stamp the edits into the part through the splice owner, which refuses a
 * result that would leave a comment range with only one half. These edits
 * rewrite `<w:p …>` open tags and the root element's attributes, so they
 * should never cut across a range marker; the owner is what makes that
 * structural rather than a reading of the scanner.
 */
const applySplices = (
  xml: string,
  edits: SpliceEdit[],
  partPath: string,
): string => {
  const patched = spliceXml(
    xml,
    edits.map(({ start, end, text }) => ({ start, end, newXml: text })),
  );
  if (patched === null) {
    throw createEnsureParaIdsError(
      `Stamping paragraph ids into ${partPath} broke a comment range`,
    );
  }
  return patched;
};

const createEnsureParaIdsError = (
  message: string,
  cause?: unknown,
): EnsureParaIdsError =>
  new EnsureParaIdsError({
    reason: ENSURE_PARA_IDS_REASONS.NORMALIZATION_FAILED,
    message,
    ...(cause === undefined ? {} : { cause }),
  });

export const collectExistingParaIds = (
  xml: string,
  into: Set<string>,
): void => {
  for (const match of xml.matchAll(ANY_PARA_ID_PATTERN)) {
    // SAFETY: named group `id` always present when the pattern matches
    const id = match.groups!["id"]!;
    if (id.length > 0) {
      into.add(id.toUpperCase());
    }
  }
};

type AttributeSpan = { start: number; end: number; value: string };
type ParagraphOpenTag = {
  nameEnd: number;
  inFallback: boolean;
  /** The innermost `mc:AlternateContent` the paragraph is drawn in, if any. */
  alternateContent: XmlElement | undefined;
  paraId: AttributeSpan | undefined;
  textId: AttributeSpan | undefined;
};

type PartSpelling = {
  xml: string;
  root: XmlElement;
  rootNameEnd: number;
  ignorable: AttributeSpan | undefined;
  paragraphs: ParagraphOpenTag[];
  w14Prefix: string;
  mcPrefix: string;
};

/** Resolve every element and attribute in its own scope; retain source offsets for splices. */
const scanPartSpelling = (xml: string, partPath: string): PartSpelling => {
  let root: XmlElement | undefined;
  let rootNameEnd = 0;
  let ignorable: AttributeSpan | undefined;
  const paragraphs: ParagraphOpenTag[] = [];
  const fallbacks = new WeakSet<XmlElement>();
  const alternateContents = new WeakMap<XmlElement, XmlElement>();
  const bindings = new Map<string, Set<string>>();
  const scanned = scanStreamingXmlElements(
    xml,
    ({ element, attributeValueSpans: spans, nameEnd, parent }) => {
      const elementName = element.name ?? "";
      const unboundElement =
        elementName.includes(":") && getNamespaceUri(element) === undefined;
      const unboundAttribute = Object.keys(element.attributes ?? {}).some(
        (attribute) =>
          attribute.includes(":") &&
          !attribute.startsWith("xmlns:") &&
          resolveAttributeNamespaceUri(element, attribute) === undefined,
      );
      if (unboundElement || unboundAttribute)
        throw new EnsureParaIdsError({
          reason: ENSURE_PARA_IDS_REASONS.NAMESPACE_INVALID,
          message: `Undeclared namespace prefix in ${partPath}`,
        });
      if (parent === undefined) {
        if (root !== undefined)
          throw createEnsureParaIdsError(`Multiple roots in ${partPath}`);
        root = element;
        rootNameEnd = nameEnd;
      }
      for (const [name, value] of Object.entries(element.attributes ?? {})) {
        if (name !== "xmlns" && !name.startsWith("xmlns:")) continue;
        const prefix = name === "xmlns" ? "" : name.slice(6);
        let uris = bindings.get(prefix);
        if (uris === undefined) {
          uris = new Set();
          bindings.set(prefix, uris);
        }
        uris.add(String(value));
      }
      const inFallback =
        (parent !== undefined && fallbacks.has(parent)) ||
        (getLocalName(element.name ?? "") === "Fallback" &&
          getNamespaceUri(element) === MC_NAMESPACE_URI);
      if (inFallback) fallbacks.add(element);
      const alternateContent =
        getLocalName(element.name ?? "") === "AlternateContent" &&
        getNamespaceUri(element) === MC_NAMESPACE_URI
          ? element
          : parent && alternateContents.get(parent);
      if (alternateContent) alternateContents.set(element, alternateContent);
      let paraId: AttributeSpan | undefined;
      let textId: AttributeSpan | undefined;
      for (const [name, span] of spans) {
        const namespace = resolveAttributeNamespaceUri(element, name);
        const localName = getLocalName(name);
        const value = String(element.attributes?.[name] ?? "");
        if (
          element === root &&
          namespace === MC_NAMESPACE_URI &&
          localName === "Ignorable"
        )
          ignorable = { ...span, value };
        if (
          namespace !== W14_NAMESPACE_URI &&
          !WORDPROCESSINGML_NAMESPACE_URIS.has(namespace ?? "")
        )
          continue;
        if (localName === "paraId") {
          if (paraId === undefined || namespace === W14_NAMESPACE_URI)
            paraId = { ...span, value };
        }
        if (localName === "textId") textId = { ...span, value };
      }
      if (
        getLocalName(element.name ?? "") === "p" &&
        WORDPROCESSINGML_NAMESPACE_URIS.has(getNamespaceUri(element) ?? "")
      )
        paragraphs.push({
          nameEnd,
          inFallback,
          alternateContent,
          paraId,
          textId,
        });
    },
  );
  if (scanned.status === "unsupported" || root === undefined) {
    // A reference XML refuses is named with its place, as every open path
    // names it.
    const malformed =
      scanned.status === "unsupported" &&
      scanned.refusal.type === "malformedReference"
        ? malformedXmlReferenceError(partPath, xml)
        : null;
    throw malformed === null
      ? createEnsureParaIdsError(`Malformed XML in ${partPath}`)
      : createEnsureParaIdsError(malformed.message, malformed);
  }
  const rootElement = root;
  const prefixFor = (uri: string, preferred: string): string => {
    for (const [prefix, boundUri] of rootElement.namespaceScope?.bindings ??
      []) {
      if (prefix === "" || boundUri !== uri) continue;
      if (uri !== W14_NAMESPACE_URI || bindings.get(prefix)?.size === 1)
        return prefix;
    }
    let candidate = preferred;
    for (let suffix = 1; bindings.has(candidate); suffix += 1)
      candidate = `${preferred}_${suffix}`;
    return candidate;
  };
  return {
    xml,
    root,
    rootNameEnd,
    ignorable,
    paragraphs,
    w14Prefix: prefixFor(W14_NAMESPACE_URI, MC_IGNORABLE_W14),
    mcPrefix: prefixFor(MC_NAMESPACE_URI, "mc"),
  };
};

export const isWordprocessingElement = (
  element: XmlElement,
  localName: string,
): boolean =>
  getLocalName(element.name ?? "") === localName &&
  WORDPROCESSINGML_NAMESPACE_URIS.has(getNamespaceUri(element) ?? "");

/**
 * The block that decides whether a cell ends in a paragraph, as the parser
 * reads it: the last paragraph or table among the cell's children, content
 * controls and custom XML elements read through.
 */
const lastFlowBlock = (
  children: readonly XmlElement[],
): "paragraph" | "table" | undefined => {
  for (let index = children.length - 1; index >= 0; index -= 1) {
    const child = children[index];
    if (child === undefined) continue;
    if (isWordprocessingElement(child, "p")) return "paragraph";
    if (isWordprocessingElement(child, "tbl")) return "table";
    if (isWordprocessingElement(child, "sdt")) {
      const content = getChildElements(child).find((element) =>
        isWordprocessingElement(element, "sdtContent"),
      );
      const nested = content && lastFlowBlock(getChildElements(content));
      if (nested !== undefined) return nested;
    }
    if (isWordprocessingElement(child, "customXml")) {
      const nested = lastFlowBlock(getChildElements(child));
      if (nested !== undefined) return nested;
    }
  }
  return undefined;
};

/** A place in a story part that holds a paragraph. */
export type StoryParagraphSlot =
  | { type: "paragraph"; element: XmlElement }
  /** The empty paragraph a cell that ends without one is read with. */
  | { type: "impliedCellParagraph"; cell: XmlElement };

/**
 * A story part's paragraphs in the order {@link ensureParaIds} numbers them:
 * every `w:p` outside `mc:Fallback` in document order, and after a cell's
 * content the paragraph the cell implies when it ends in a table or holds
 * none. A fallback branch is another rendering of content read elsewhere.
 */
export const storyParagraphSlots = (root: XmlElement): StoryParagraphSlot[] => {
  const slots: StoryParagraphSlot[] = [];
  const visit = (element: XmlElement): void => {
    if (
      getLocalName(element.name ?? "") === "Fallback" &&
      getNamespaceUri(element) === MC_NAMESPACE_URI
    )
      return;
    if (isWordprocessingElement(element, "p"))
      slots.push({ type: "paragraph", element });
    const children = getChildElements(element);
    for (const child of children) visit(child);
    if (
      isWordprocessingElement(element, "tc") &&
      lastFlowBlock(children) !== "paragraph"
    )
      slots.push({ type: "impliedCellParagraph", cell: element });
  };
  visit(root);
  return slots;
};

/**
 * A cell must end in a paragraph. One that ends in a table, or holds none,
 * is read with the empty paragraph it implies there; that paragraph is
 * written into the part so it can carry an id as every other paragraph does.
 */
const withImpliedCellParagraphs = (xml: string): string => {
  if (!/<(?:[\w.-]+:)?tc[\s/>]/u.test(xml)) return xml;
  const parsed = parseStreamingXmlWithSourceRanges(xml);
  if (parsed.status !== "parsed") return xml;
  const edits: { start: number; end: number; text: string }[] = [];
  for (const slot of storyParagraphSlots(parsed.value)) {
    if (slot.type !== "impliedCellParagraph") continue;
    const { cell } = slot;
    const range = getXmlSourceRange(cell);
    const name = cell.name ?? "";
    const separator = name.indexOf(":");
    const paragraph = `<${separator < 0 ? "" : name.slice(0, separator + 1)}p/>`;
    if (range !== undefined && xml.slice(range.end - 2, range.end) === "/>")
      edits.push({
        start: range.start,
        end: range.end,
        text: `${xml.slice(range.start, range.end - 2).trimEnd()}>${paragraph}</${name}>`,
      });
    else if (range !== undefined) {
      const close = xml.lastIndexOf("</", range.end - 1);
      if (close > range.start)
        edits.push({ start: close, end: close, text: paragraph });
    }
  }
  if (edits.length === 0) return xml;
  return (
    spliceXml(
      xml,
      edits.map(({ start, end, text }) => ({ start, end, newXml: text })),
    ) ?? xml
  );
};

const partSpelling = (xml: string, partPath: string): PartSpelling => {
  const scanned = Result.try({
    try: () => scanPartSpelling(xml, partPath),
    catch: (cause) => cause,
  });
  if (scanned.isOk()) return scanned.value;
  const error = scanned.error;
  if (
    error instanceof EnsureParaIdsError &&
    error.reason === ENSURE_PARA_IDS_REASONS.NAMESPACE_INVALID
  ) {
    const repaired = repairMissingDrawingNamespace(xml);
    if (repaired !== xml) return scanPartSpelling(repaired, partPath);
  }
  throw error;
};

/**
 * The ids `mc:Fallback` paragraphs hold. A fallback renders its
 * `mc:Choice` again, and a producer that regenerates it repeats the choice's
 * paragraph ids, so an id the choice of the same `mc:AlternateContent` holds
 * is the choice paragraph's own, not a fallback id it repeats.
 */
const collectFallbackParaIds = (
  spelling: PartSpelling,
  into: Set<string>,
): void => {
  const choiceIds = new Map<XmlElement, Set<string>>();
  for (const { inFallback, alternateContent, paraId } of spelling.paragraphs) {
    if (inFallback || !alternateContent || !paraId?.value) continue;
    let ids = choiceIds.get(alternateContent);
    if (ids === undefined) {
      ids = new Set();
      choiceIds.set(alternateContent, ids);
    }
    ids.add(paraId.value.toUpperCase());
  }
  for (const { inFallback, alternateContent, paraId } of spelling.paragraphs) {
    if (!inFallback || !paraId?.value) continue;
    const id = paraId.value.toUpperCase();
    if (alternateContent && choiceIds.get(alternateContent)?.has(id)) continue;
    into.add(id);
  }
};

export type MintContext = {
  /** Every id in the document (all parts) plus every id minted so far. */
  taken: Set<string>;
};

/**
 * What a fresh id is keyed on for each paragraph of a part, by ordinal (from
 * 1): its own markup, and how many earlier paragraphs repeat that markup. An
 * edit elsewhere in the part changes neither, so a paragraph it left as it
 * was is given the same id in both versions of the document.
 */
export const paragraphMintKeys = (
  xml: string,
): readonly string[] | undefined => {
  const parsed = parseStreamingXmlWithSourceRanges(xml);
  if (parsed.status !== "parsed") return undefined;
  const repeats = new Map<string, number>();
  const keys: string[] = [];
  for (const slot of storyParagraphSlots(parsed.value)) {
    if (slot.type !== "paragraph") return undefined;
    const range = getXmlSourceRange(slot.element);
    if (range === undefined) return undefined;
    const markup = xml.slice(range.start, range.end);
    const repeat = repeats.get(markup) ?? 0;
    repeats.set(markup, repeat + 1);
    keys.push(`${repeat}:${markup}`);
  }
  return keys;
};

/**
 * Deterministic fresh id for the paragraph at `ordinal` in `partPath`,
 * salted past collisions with any id anywhere in the document. `keys` are
 * the part's {@link paragraphMintKeys}; a part they cannot be read for is
 * keyed on the ordinal alone.
 */
export const mintParaId = (
  context: MintContext,
  partPath: string,
  ordinal: number,
  keys: readonly string[] | undefined,
): string => {
  const key = keys?.[ordinal - 1] ?? String(ordinal);
  const seed = `${partPath}:${key}`;
  let id = deterministicHexId(seed);
  for (let salt = 1; context.taken.has(id); salt += 1) {
    id = deterministicHexId(`${seed}:${salt}`);
  }
  context.taken.add(id);
  return id;
};

type PartScanResult = {
  edits: SpliceEdit[];
  minted: string[];
  /**
   * The id minted for each paragraph the package names no usable id for, by
   * its ordinal among the part's {@link storyParagraphSlots} (from 1).
   */
  unnamed: Map<number, string>;
  /** The part's paragraphs outside `mc:Fallback`, implied ones included. */
  paragraphs: number;
  assigned: number;
  deduplicated: number;
};

/**
 * One forward scan over a part: every paragraph open tag outside
 * `mc:Fallback` either keeps its paraId (valid, first occurrence) or gets a
 * splice edit minting / replacing one. `seen` spans parts so the
 * first-occurrence rule is document-wide across the scan order.
 */
type ScanPartOptions = {
  partPath: string;
  spelling: PartSpelling;
  context: MintContext;
  seen: Set<string>;
};

const scanPart = ({
  partPath,
  spelling,
  context,
  seen,
}: ScanPartOptions): PartScanResult => {
  const edits: SpliceEdit[] = [];
  const minted: string[] = [];
  const unnamed = new Map<number, string>();
  const w14 = spelling.w14Prefix;
  // Read once, at the first paragraph that takes a fresh id.
  let keys: readonly string[] | undefined | null = null;
  const mint = (ordinal: number): string => {
    keys ??= paragraphMintKeys(spelling.xml);
    return mintParaId(context, partPath, ordinal, keys);
  };
  let assigned = 0;
  let deduplicated = 0;
  let ordinal = 0;

  for (const { nameEnd, inFallback, paraId, textId } of spelling.paragraphs) {
    if (inFallback) {
      continue;
    }
    ordinal += 1;

    if (paraId === undefined) {
      const id = mint(ordinal);
      if (textId !== undefined) {
        edits.push({
          start: nameEnd,
          end: nameEnd,
          text: ` ${w14}:paraId="${id}"`,
        });
        edits.push({ start: textId.start, end: textId.end, text: id });
      } else {
        edits.push({
          start: nameEnd,
          end: nameEnd,
          text: ` ${w14}:paraId="${id}" ${w14}:textId="${id}"`,
        });
      }
      assigned += 1;
      minted.push(id);
      unnamed.set(ordinal, id);
      seen.add(id);
      continue;
    }

    const value = paraId.value.toUpperCase();
    // The parser reads anything but eight hex digits as no id at all.
    const unassigned =
      !isValidHexId(value) || RESERVED_ZERO_ID_PATTERN.test(value);
    if (!unassigned && !seen.has(value)) {
      seen.add(value);
      continue;
    }

    const id = mint(ordinal);
    edits.push({ start: paraId.start, end: paraId.end, text: id });
    if (textId !== undefined) {
      edits.push({ start: textId.start, end: textId.end, text: id });
    } else {
      edits.push({
        start: nameEnd,
        end: nameEnd,
        text: ` ${w14}:textId="${id}"`,
      });
    }
    if (unassigned) {
      assigned += 1;
      unnamed.set(ordinal, id);
    } else {
      deduplicated += 1;
    }
    minted.push(id);
    seen.add(id);
  }

  return {
    edits,
    minted,
    unnamed,
    paragraphs: ordinal,
    assigned,
    deduplicated,
  };
};

/**
 * Splice edits ensuring the part's root element declares `xmlns:w14` /
 * `xmlns:mc` and lists `w14` in `mc:Ignorable`, so consumers that predate the
 * 2010 extensions skip the new attributes instead of rejecting the part.
 */
const ensureRootNamespaces = ({
  root,
  rootNameEnd,
  ignorable,
  w14Prefix,
  mcPrefix,
}: PartSpelling): SpliceEdit[] => {
  const edits: SpliceEdit[] = [];
  const declarations: string[] = [];
  if (resolveNamespaceUri(root.namespaceScope, mcPrefix) !== MC_NAMESPACE_URI)
    declarations.push(` xmlns:${mcPrefix}="${MC_NAMESPACE_URI}"`);
  if (resolveNamespaceUri(root.namespaceScope, w14Prefix) !== W14_NAMESPACE_URI)
    declarations.push(` xmlns:${w14Prefix}="${W14_NAMESPACE_URI}"`);
  if (ignorable !== undefined) {
    const tokens = ignorable.value
      .split(/\s+/u)
      .filter((token) => token.length > 0);
    if (!tokens.includes(w14Prefix))
      edits.push({
        start: ignorable.end,
        end: ignorable.end,
        text: tokens.length === 0 ? w14Prefix : ` ${w14Prefix}`,
      });
  } else {
    declarations.push(` ${mcPrefix}:Ignorable="${w14Prefix}"`);
  }
  if (declarations.length > 0)
    edits.push({
      start: rootNameEnd,
      end: rootNameEnd,
      text: declarations.join(""),
    });
  return edits;
};

const mainDocumentPart = (rels: string, entries: readonly string[]): string => {
  const root = parseXmlDocument(rels);
  if (
    root === null ||
    getLocalName(root.name ?? "") !== "Relationships" ||
    getNamespaceUri(root) !== PACKAGE_RELATIONSHIP_URI
  )
    throw createEnsureParaIdsError("Malformed package relationships");
  const documents = getChildElements(root).filter((element) => {
    if (
      getLocalName(element.name ?? "") !== "Relationship" ||
      getNamespaceUri(element) !== PACKAGE_RELATIONSHIP_URI
    )
      return false;
    const type = getAttribute(element, null, "Type");
    return [...OFFICE_RELATIONSHIP_NAMESPACE_URIS].some(
      (uri) => type === `${uri}/officeDocument`,
    );
  });
  const document = documents.at(0);
  if (
    documents.length !== 1 ||
    document === undefined ||
    getAttribute(document, null, "TargetMode") === "External"
  )
    throw createEnsureParaIdsError(
      "Package must have one internal officeDocument relationship",
    );
  const target = getAttribute(document, null, "Target");
  const partPath =
    target === null
      ? undefined
      : resolvePackageRelationshipTarget(target, "_rels/.rels");
  if (partPath === undefined)
    throw createEnsureParaIdsError("Invalid officeDocument target");
  const name = entries.find((entry) => entry.toLowerCase() === partPath);
  if (name === undefined)
    throw createEnsureParaIdsError("officeDocument part not found");
  return name;
};

/** The `.rels` part that holds `partName`'s relationships. */
const relationshipsPathFor = (partName: string): string => {
  const slash = partName.lastIndexOf("/");
  return `${partName.slice(0, slash + 1)}_rels/${partName.slice(slash + 1)}.rels`;
};

/**
 * The header, footer, and note parts the main part's relationships name.
 * The parser follows these relationships, so a story at a custom path
 * (`word/custom/defaultHeader.xml`) is as editable as `word/header1.xml`.
 */
const relatedStoryParts = (
  rels: string | null,
  relsPath: string,
  entries: readonly string[],
): ReadonlySet<string> => {
  const parts = new Set<string>();
  const root = rels === null ? null : parseXmlDocument(rels);
  if (
    root === null ||
    getLocalName(root.name ?? "") !== "Relationships" ||
    getNamespaceUri(root) !== PACKAGE_RELATIONSHIP_URI
  )
    return parts;
  const storyTypes = new Set<string>();
  for (const uri of OFFICE_RELATIONSHIP_NAMESPACE_URIS)
    for (const type of STORY_RELATIONSHIP_TYPES)
      storyTypes.add(`${uri}/${type}`);
  for (const element of getChildElements(root)) {
    if (
      getLocalName(element.name ?? "") !== "Relationship" ||
      getNamespaceUri(element) !== PACKAGE_RELATIONSHIP_URI ||
      !storyTypes.has(getAttribute(element, null, "Type") ?? "") ||
      getAttribute(element, null, "TargetMode") === "External"
    )
      continue;
    const target = getAttribute(element, null, "Target");
    const partPath =
      target === null
        ? undefined
        : resolvePackageRelationshipTarget(target, relsPath);
    const name =
      partPath === undefined
        ? undefined
        : entries.find((entry) => entry.toLowerCase() === partPath);
    if (name !== undefined) parts.add(name);
  }
  return parts;
};

/** OPC part names are case-insensitive; compare lowercased. */
const matchesTargetPattern = (path: string): boolean =>
  TARGET_PART_PATTERNS.some((pattern) => pattern.test(path.toLowerCase()));

const toUint8Array = (docx: Uint8Array | ArrayBuffer): Uint8Array =>
  docx instanceof Uint8Array ? docx : new Uint8Array(docx);

const hasDigitalSignatureParts = (zip: ZipArchive): boolean =>
  zip
    .entries()
    .some(({ path }) =>
      path.toLowerCase().startsWith(DIGITAL_SIGNATURE_PART_PREFIX),
    );

type PackageStoryPartsOptions = {
  documentPartName: string;
  /** The main part's relationships part, when the package has one. */
  documentRels: string | null;
  entries: readonly string[];
};

type PackageStoryParts = {
  storyParts: ReadonlySet<string>;
  isTargetPart: (name: string) => boolean;
};

/** The parts besides the main one whose paragraphs get ids. */
export const packageStoryParts = ({
  documentPartName,
  documentRels,
  entries,
}: PackageStoryPartsOptions): PackageStoryParts => {
  const storyParts = relatedStoryParts(
    documentRels,
    relationshipsPathFor(documentPartName),
    entries,
  );
  return {
    storyParts,
    isTargetPart: (name) =>
      name !== documentPartName &&
      (storyParts.has(name) || matchesTargetPattern(name)),
  };
};

/** What {@link assignPackageParaIds} reads: a package's XML parts by name. */
type AssignPackageParaIdsOptions = {
  partTexts: ReadonlyMap<string, string>;
  documentPartName: string;
  /** Whether a part other than the main one holds stories whose paragraphs get ids. */
  isTargetPart: (name: string) => boolean;
};

/** One story part as it scanned, and what the scan gave its paragraphs. */
type PartParaIdScan = {
  /** The part as it scanned, with the empty paragraphs its cells imply. */
  spelling: PartSpelling;
  scan: PartScanResult;
};

/**
 * The paragraph ids a package's story parts take: the one assignment behind
 * {@link ensureParaIds}, which writes them, and {@link assignPackageParaIds},
 * which names them. Story parts in the order their paragraphs are numbered.
 */
const scanPackageParaIds = ({
  partTexts,
  documentPartName,
  isTargetPart,
}: AssignPackageParaIdsOptions): ReadonlyMap<string, PartParaIdScan> => {
  const taken = new Set<string>();
  for (const text of partTexts.values()) {
    collectExistingParaIds(text, taken);
  }

  if (!partTexts.has(documentPartName))
    throw createEnsureParaIdsError("officeDocument part is unreadable");
  const context: MintContext = { taken };

  // document.xml first so its paragraphs win the first-occurrence rule, then
  // the auxiliary parts in a deterministic order.
  const targetParts = [
    documentPartName,
    ...[...partTexts.keys()]
      .filter((name) => isTargetPart(name))
      .sort((a, b) => a.localeCompare(b)),
  ];

  // IDs in parts or fallback branches that this pass deliberately leaves
  // untouched own their values. Seed duplicate resolution with them so an
  // editable paragraph cannot preserve a conflicting ID.
  const spellings = new Map<string, PartSpelling>();
  const seen = new Set<string>();
  for (const [partPath, xml] of partTexts) {
    if (partPath === documentPartName || isTargetPart(partPath)) {
      const spelling = partSpelling(withImpliedCellParagraphs(xml), partPath);
      spellings.set(partPath, spelling);
      collectFallbackParaIds(spelling, seen);
    } else {
      collectExistingParaIds(xml, seen);
    }
  }

  const parts = new Map<string, PartParaIdScan>();
  for (const partPath of targetParts) {
    // SAFETY: every target part was given a spelling above
    const spelling = spellings.get(partPath)!;
    parts.set(partPath, {
      spelling,
      scan: scanPart({ partPath, spelling, context, seen }),
    });
  }
  return parts;
};

/** The ids a story part's paragraphs take where the package names none. */
export type UnnamedParagraphIds = {
  /** Each such paragraph's id, by its ordinal among the part's {@link storyParagraphSlots} (from 1). */
  byOrdinal: ReadonlyMap<number, string>;
  /** How many slots the part has, as the assignment counted them. */
  paragraphs: number;
};

/**
 * The id {@link ensureParaIds} writes for each story paragraph the package
 * names no usable id for, by part, without writing it: a package opened as
 * bytes on either engine so gives each paragraph the same id. A fresh id is
 * keyed on the main part's content, the part, and the paragraph's ordinal in
 * it. A repeated authored id is left for its readers to judge.
 */
export const assignPackageParaIds = (
  options: AssignPackageParaIdsOptions,
): ReadonlyMap<string, UnnamedParagraphIds> => {
  const assigned = new Map<string, UnnamedParagraphIds>();
  for (const [partPath, { scan }] of scanPackageParaIds(options))
    assigned.set(partPath, {
      byOrdinal: scan.unnamed,
      paragraphs: scan.paragraphs,
    });
  return assigned;
};

type PackageParaIdUpdates = {
  updates: ReadonlyMap<string, string>;
  mintedParaIds: readonly string[];
  assigned: number;
  deduplicated: number;
};

/** The part texts that write an assignment's ids, and the ids written. */
const packageParaIdUpdates = (
  partTexts: ReadonlyMap<string, string>,
  parts: ReadonlyMap<string, PartParaIdScan>,
): PackageParaIdUpdates => {
  const updates = new Map<string, string>();
  const mintedParaIds: string[] = [];
  let assigned = 0;
  let deduplicated = 0;
  for (const [partPath, { spelling, scan }] of parts) {
    if (scan.edits.length === 0) {
      if (spelling.xml !== partTexts.get(partPath))
        updates.set(partPath, spelling.xml);
      continue;
    }
    assigned += scan.assigned;
    deduplicated += scan.deduplicated;
    // One push per id: spreading a large part's ids as arguments can exceed
    // the engine's argument limit.
    for (const id of scan.minted) {
      mintedParaIds.push(id);
    }
    updates.set(
      partPath,
      applySplices(
        spelling.xml,
        [...scan.edits, ...ensureRootNamespaces(spelling)],
        partPath,
      ),
    );
  }
  return { updates, mintedParaIds, assigned, deduplicated };
};

/**
 * Backfill `w14:paraId` on every paragraph of a `.docx` buffer. See the
 * module doc for the exact contract. Throws {@link EnsureParaIdsError} when
 * the buffer is not a WordprocessingML package or a part is malformed.
 */
const ensureParaIdsInternal = async (
  docx: Uint8Array | ArrayBuffer,
  options: EnsureParaIdsOptions,
): Promise<EnsureParaIdsResult> => {
  // Bounded read: entry count, per-entry size, and cumulative uncompressed
  // size are all capped before any part's XML is materialized as a string.
  // `docx` here is untrusted ingest input (see the module doc comment), so
  // an attacker-crafted archive with thousands of parts or a decompression-
  // bomb entry must fail fast instead of exhausting memory.
  const archive = await loadDocxArchive(docx);

  const relsPath = archive.entries.find(
    (name) => name.toLowerCase() === "_rels/.rels",
  );
  const rels =
    relsPath === undefined ? null : await archive.readEntryString(relsPath);
  if (rels === null)
    throw createEnsureParaIdsError("Package relationships not found");
  const documentPartName = mainDocumentPart(rels, archive.entries);
  const documentRelsPath = relationshipsPathFor(documentPartName);
  const documentRelsName = archive.entries.find(
    (name) => name.toLowerCase() === documentRelsPath.toLowerCase(),
  );
  const { storyParts, isTargetPart } = packageStoryParts({
    documentPartName,
    documentRels:
      documentRelsName === undefined
        ? null
        : await archive.readEntryString(documentRelsName),
    entries: archive.entries,
  });
  const xmlPartNames = archive.entries.filter(
    (name) =>
      name.toLowerCase().endsWith(".xml") ||
      name === documentPartName ||
      storyParts.has(name),
  );

  const partTexts = new Map<string, string>();
  for (const name of xmlPartNames) {
    // oxlint-disable-next-line no-await-in-loop -- archive.readEntryString serializes reads to enforce a shared cumulative byte budget across parts
    const text = await archive.readEntryString(name);
    if (text !== null) {
      partTexts.set(name, text);
    }
  }

  const { updates, mintedParaIds, assigned, deduplicated } =
    packageParaIdUpdates(
      partTexts,
      scanPackageParaIds({ partTexts, documentPartName, isTargetPart }),
    );

  if (updates.size === 0) {
    return {
      docx: toUint8Array(docx),
      assigned: 0,
      deduplicated: 0,
      alreadyComplete: true,
      mintedParaIds: [],
    };
  }

  // Only the write-back path needs a raw ZipArchive (signature detection, writing
  // new part bytes, and repackaging). `docx` was already validated against
  // the entry-count/size caps above, so this second parse runs on a
  // buffer whose shape is already bounded.
  const zip = await ZipArchive.read(docx);

  if (
    hasDigitalSignatureParts(zip) &&
    options.allowSignedPackageMutation !== true
  ) {
    throw new EnsureParaIdsError({
      reason: ENSURE_PARA_IDS_REASONS.SIGNED_PACKAGE,
      message:
        "Refusing to normalize a digitally signed package because rewriting OOXML invalidates its signatures. Warn the user and pass allowSignedPackageMutation only if invalidation is acceptable.",
    });
  }

  for (const [partPath, content] of updates) {
    const sourceEntry = zip.entry(partPath);
    if (sourceEntry === undefined) {
      throw createEnsureParaIdsError(
        `Package part disappeared during normalization: ${partPath}`,
      );
    }
    zip.set(partPath, content, { modified: sourceEntry.modified });
  }
  const output = await zip.generate({ level: 6 });

  return {
    docx: output,
    assigned,
    deduplicated,
    alreadyComplete: false,
    mintedParaIds,
  };
};

/**
 * Backfill `w14:paraId` on every paragraph of a `.docx` buffer. All package,
 * XML, and compression failures are surfaced as {@link EnsureParaIdsError}.
 */
export const ensureParaIds = async (
  docx: Uint8Array | ArrayBuffer,
  options: EnsureParaIdsOptions = {},
): Promise<EnsureParaIdsResult> => {
  try {
    return await ensureParaIdsInternal(docx, options);
  } catch (error) {
    if (error instanceof EnsureParaIdsError) {
      throw error;
    }
    const message = error instanceof Error ? error.message : String(error);
    throw createEnsureParaIdsError(
      `Failed to normalize paragraph IDs: ${message}`,
      error,
    );
  }
};
