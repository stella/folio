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
 * - Parts covered: the officeDocument relationship target, `word/header*.xml`, `word/footer*.xml`,
 *   `word/footnotes.xml`, `word/endnotes.xml`. Paragraphs nested in table
 *   cells and in `mc:Choice` text boxes are plain `<w:p>` elements inside
 *   those parts and are covered by the same scan. The comments part mints its
 *   own deterministic paraIds at save time (see `commentSerializer`) and is
 *   left untouched; its ids still count toward uniqueness.
 * - Paragraphs inside `mc:Fallback` are never modified: Word regenerates the
 *   fallback branch (duplicating the `mc:Choice` ids) on save, so stamping or
 *   deduplicating there would churn on every Word round-trip.
 * - Existing ids are preserved. Only three cases get a fresh id: a missing
 *   `paraId`, the reserved all-zero value (Word reads `00000000` as "no id"),
 *   and a duplicate of an id already seen earlier in the scan (first
 *   occurrence keeps it — the same rule `ParaIdAllocatorExtension` applies in
 *   the editor). `w14:textId` is written alongside a newly minted `paraId`
 *   (same value) and never touched otherwise; it is a text-revision marker,
 *   not identity.
 * - Fresh ids are deterministic (`deterministicHexId` over document content,
 *   part path, and paragraph ordinal), so the pass is a pure function of the
 *   input bytes: retrying an ingest produces identical output.
 * - Each patched part's root element gets `xmlns:w14` / `xmlns:mc`
 *   declarations and a `mc:Ignorable` listing `w14` when missing — non-Word
 *   producers declare neither, and absent `mc:Ignorable` handling is what
 *   makes pre-2010 consumers choke on the new attributes.
 * - Prefixes are aliases: elements and attributes resolve against their
 *   in-scope namespace URI, including Strict and Transitional profiles and
 *   nested rebinding. Minted attributes use a root prefix with no conflicting
 *   nested binding; a fresh prefix is declared when necessary.
 * - Idempotent: a document that already has full coverage is returned as the
 *   original bytes, untouched (`alreadyComplete: true`). A body with
 *   paragraphs the scan did not see is refused, never reported complete.
 * - Digitally signed packages are returned untouched when already complete.
 *   When normalization would rewrite the package, it fails unless the caller
 *   explicitly allows signature invalidation after warning the user.
 */
import { TaggedError } from "better-result";
import JSZip from "jszip";

import { deterministicHexId } from "../utils/hexId";
import { resolvePackageRelationshipTarget } from "./packageParts";
import { spliceXml } from "./selectiveXmlPatch";
import { loadDocxArchive } from "./server/boundedArchive";
import { scanStreamingXmlElements } from "./streamingXmlParser";
import { resolveNamespaceUri } from "./xmlNamespaceContext";
import {
  getAttribute,
  getChildElements,
  getLocalName,
  getNamespaceUri,
  parseXmlDocument,
  resolveAttributeNamespaceUri,
  WORDPROCESSINGML_NAMESPACE_URIS,
  OFFICE_RELATIONSHIP_NAMESPACE_URIS,
  type XmlElement,
} from "./xmlParser";

export const ENSURE_PARA_IDS_REASONS = {
  NORMALIZATION_FAILED: "normalization-failed",
  NAMESPACE_INVALID: "namespace-invalid",
  SIGNED_PACKAGE: "signed-package",
} as const;

/** A malformed or unsupported package prevented paragraph-ID normalization. */
export class EnsureParaIdsError extends TaggedError("EnsureParaIdsError")<{
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

const W14_NAMESPACE_URI = "http://schemas.microsoft.com/office/word/2010/wordml";
const MC_NAMESPACE_URI = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const MC_IGNORABLE_W14 = "w14";
const DIGITAL_SIGNATURE_PART_PREFIX = "_xmlsignatures/";
/** Word reads an all-zero `w14:paraId` as "no id assigned". */
const RESERVED_ZERO_ID_PATTERN = /^0*$/u;

const PACKAGE_RELATIONSHIP_URI = "http://schemas.openxmlformats.org/package/2006/relationships";
const TARGET_PART_PATTERNS = [
  /^word\/header\d*\.xml$/u,
  /^word\/footer\d*\.xml$/u,
  /^word\/footnotes\.xml$/u,
  /^word\/endnotes\.xml$/u,
];

/** Reserve package ids even in parts whose content the pass leaves untouched. */
const ANY_PARA_ID_PATTERN = /\s[^\s=<>/"':]+:paraId\s*=\s*(?<quote>["'])(?<id>[\s\S]*?)\k<quote>/gu;

type SpliceEdit = { start: number; end: number; text: string };

/**
 * Stamp the edits into the part through the splice owner, which refuses a
 * result that would leave a comment range with only one half. These edits
 * rewrite `<w:p …>` open tags and the root element's attributes, so they
 * should never cut across a range marker; the owner is what makes that
 * structural rather than a reading of the scanner.
 */
const applySplices = (xml: string, edits: SpliceEdit[], partPath: string): string => {
  const patched = spliceXml(
    xml,
    edits.map(({ start, end, text }) => ({ start, end, newXml: text })),
  );
  if (patched === null) {
    throw createEnsureParaIdsError(`Stamping paragraph ids into ${partPath} broke a comment range`);
  }
  return patched;
};

const createEnsureParaIdsError = (message: string, cause?: unknown): EnsureParaIdsError =>
  new EnsureParaIdsError({
    reason: ENSURE_PARA_IDS_REASONS.NORMALIZATION_FAILED,
    message,
    ...(cause === undefined ? {} : { cause }),
  });

const collectExistingParaIds = (xml: string, into: Set<string>): void => {
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
  paraId: AttributeSpan | undefined;
  textId: AttributeSpan | undefined;
};

type PartSpelling = {
  root: XmlElement;
  rootNameEnd: number;
  ignorable: AttributeSpan | undefined;
  paragraphs: ParagraphOpenTag[];
  w14Prefix: string;
  mcPrefix: string;
};

/** Resolve every element and attribute in its own scope; retain source offsets for splices. */
const partSpelling = (xml: string, partPath: string): PartSpelling => {
  let root: XmlElement | undefined;
  let rootNameEnd = 0;
  let ignorable: AttributeSpan | undefined;
  const paragraphs: ParagraphOpenTag[] = [];
  const fallbacks = new WeakSet<XmlElement>();
  const bindings = new Map<string, Set<string>>();
  const scanned = scanStreamingXmlElements(
    xml,
    ({ element, attributeValueSpans: spans, nameEnd, parent }) => {
      const elementName = element.name ?? "";
      const unboundElement = elementName.includes(":") && getNamespaceUri(element) === undefined;
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
        if (root !== undefined) throw createEnsureParaIdsError(`Multiple roots in ${partPath}`);
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
      let paraId: AttributeSpan | undefined;
      let textId: AttributeSpan | undefined;
      for (const [name, span] of spans) {
        const namespace = resolveAttributeNamespaceUri(element, name);
        const localName = getLocalName(name);
        const value = String(element.attributes?.[name] ?? "");
        if (element === root && namespace === MC_NAMESPACE_URI && localName === "Ignorable")
          ignorable = { ...span, value };
        if (
          namespace !== W14_NAMESPACE_URI &&
          !WORDPROCESSINGML_NAMESPACE_URIS.has(namespace ?? "")
        )
          continue;
        if (localName === "paraId") {
          if (paraId === undefined || namespace === W14_NAMESPACE_URI) paraId = { ...span, value };
        }
        if (localName === "textId") textId = { ...span, value };
      }
      if (
        getLocalName(element.name ?? "") === "p" &&
        WORDPROCESSINGML_NAMESPACE_URIS.has(getNamespaceUri(element) ?? "")
      )
        paragraphs.push({ nameEnd, inFallback, paraId, textId });
    },
  );
  if (scanned.status === "unsupported" || root === undefined)
    throw createEnsureParaIdsError(`Malformed XML in ${partPath}`);
  const rootElement = root;
  const prefixFor = (uri: string, preferred: string): string => {
    for (const [prefix, boundUri] of rootElement.namespaceScope?.bindings ?? []) {
      if (prefix === "" || boundUri !== uri) continue;
      if (uri !== W14_NAMESPACE_URI || bindings.get(prefix)?.size === 1) return prefix;
    }
    let candidate = preferred;
    for (let suffix = 1; bindings.has(candidate); suffix += 1) candidate = `${preferred}_${suffix}`;
    return candidate;
  };
  return {
    root,
    rootNameEnd,
    ignorable,
    paragraphs,
    w14Prefix: prefixFor(W14_NAMESPACE_URI, MC_IGNORABLE_W14),
    mcPrefix: prefixFor(MC_NAMESPACE_URI, "mc"),
  };
};

const collectFallbackParaIds = (spelling: PartSpelling, into: Set<string>): void => {
  for (const { inFallback, paraId } of spelling.paragraphs) {
    if (inFallback && paraId?.value) into.add(paraId.value.toUpperCase());
  }
};

type MintContext = {
  /** Every id in the document (all parts) plus every id minted so far. */
  taken: Set<string>;
  /** `deterministicHexId` fingerprint of the original document.xml. */
  docKey: string;
};

/**
 * Deterministic fresh id for the paragraph at `ordinal` in `partPath`,
 * salted past collisions with any id anywhere in the document.
 */
const mintParaId = (context: MintContext, partPath: string, ordinal: number): string => {
  const seed = `${context.docKey}:${partPath}:${ordinal}`;
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

const scanPart = ({ partPath, spelling, context, seen }: ScanPartOptions): PartScanResult => {
  const edits: SpliceEdit[] = [];
  const minted: string[] = [];
  const w14 = spelling.w14Prefix;
  let assigned = 0;
  let deduplicated = 0;
  let ordinal = 0;

  for (const { nameEnd, inFallback, paraId, textId } of spelling.paragraphs) {
    if (inFallback) {
      continue;
    }
    ordinal += 1;

    if (paraId === undefined) {
      const id = mintParaId(context, partPath, ordinal);
      if (textId !== undefined) {
        edits.push({ start: nameEnd, end: nameEnd, text: ` ${w14}:paraId="${id}"` });
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
      seen.add(id);
      continue;
    }

    const value = paraId.value.toUpperCase();
    const unassigned = RESERVED_ZERO_ID_PATTERN.test(value);
    if (!unassigned && !seen.has(value)) {
      seen.add(value);
      continue;
    }

    const id = mintParaId(context, partPath, ordinal);
    edits.push({ start: paraId.start, end: paraId.end, text: id });
    if (textId !== undefined) {
      edits.push({ start: textId.start, end: textId.end, text: id });
    } else {
      edits.push({ start: nameEnd, end: nameEnd, text: ` ${w14}:textId="${id}"` });
    }
    if (unassigned) {
      assigned += 1;
    } else {
      deduplicated += 1;
    }
    minted.push(id);
    seen.add(id);
  }

  return { edits, minted, assigned, deduplicated };
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
    const tokens = ignorable.value.split(/\s+/u).filter((token) => token.length > 0);
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
    edits.push({ start: rootNameEnd, end: rootNameEnd, text: declarations.join("") });
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
    return [...OFFICE_RELATIONSHIP_NAMESPACE_URIS].some((uri) => type === `${uri}/officeDocument`);
  });
  const document = documents.at(0);
  if (
    documents.length !== 1 ||
    document === undefined ||
    getAttribute(document, null, "TargetMode") === "External"
  )
    throw createEnsureParaIdsError("Package must have one internal officeDocument relationship");
  const target = getAttribute(document, null, "Target");
  const partPath =
    target === null ? undefined : resolvePackageRelationshipTarget(target, "_rels/.rels");
  if (partPath === undefined) throw createEnsureParaIdsError("Invalid officeDocument target");
  const name = entries.find((entry) => entry.toLowerCase() === partPath);
  if (name === undefined) throw createEnsureParaIdsError("officeDocument part not found");
  return name;
};

/** OPC part names are case-insensitive; compare lowercased. */
const isTargetPart = (path: string): boolean =>
  TARGET_PART_PATTERNS.some((pattern) => pattern.test(path.toLowerCase()));

const toUint8Array = (docx: Uint8Array | ArrayBuffer): Uint8Array =>
  docx instanceof Uint8Array ? docx : new Uint8Array(docx);

const hasDigitalSignatureParts = (zip: JSZip): boolean =>
  Object.values(zip.files).some(
    ({ dir, name }) => !dir && name.toLowerCase().startsWith(DIGITAL_SIGNATURE_PART_PREFIX),
  );

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

  const relsPath = archive.entries.find((name) => name.toLowerCase() === "_rels/.rels");
  const rels = relsPath === undefined ? null : await archive.readEntryString(relsPath);
  if (rels === null) throw createEnsureParaIdsError("Package relationships not found");
  const documentPartName = mainDocumentPart(rels, archive.entries);
  const xmlPartNames = archive.entries.filter(
    (name) => name.toLowerCase().endsWith(".xml") || name === documentPartName,
  );

  const partTexts = new Map<string, string>();
  for (const name of xmlPartNames) {
    // oxlint-disable-next-line no-await-in-loop -- archive.readEntryString serializes reads to enforce a shared cumulative byte budget across parts
    const text = await archive.readEntryString(name);
    if (text !== null) {
      partTexts.set(name, text);
    }
  }

  const taken = new Set<string>();
  for (const text of partTexts.values()) {
    collectExistingParaIds(text, taken);
  }

  const documentXml = partTexts.get(documentPartName);
  if (documentXml === undefined)
    throw createEnsureParaIdsError("officeDocument part is unreadable");
  const context: MintContext = { taken, docKey: deterministicHexId(documentXml) };

  // document.xml first so its paragraphs win the first-occurrence rule, then
  // the auxiliary parts in a deterministic order.
  const targetParts = [
    documentPartName,
    ...xmlPartNames
      .filter((name) => name !== documentPartName && isTargetPart(name))
      .sort((a, b) => a.localeCompare(b)),
  ];

  // IDs in parts or fallback branches that this pass deliberately leaves
  // untouched own their values. Seed duplicate resolution with them so an
  // editable paragraph cannot preserve a conflicting ID.
  const spellings = new Map<string, PartSpelling>();
  const seen = new Set<string>();
  for (const [partPath, xml] of partTexts) {
    if (partPath === documentPartName || isTargetPart(partPath)) {
      const spelling = partSpelling(xml, partPath);
      spellings.set(partPath, spelling);
      collectFallbackParaIds(spelling, seen);
    } else {
      collectExistingParaIds(xml, seen);
    }
  }

  const updates = new Map<string, string>();
  const mintedParaIds: string[] = [];
  let assigned = 0;
  let deduplicated = 0;

  for (const partPath of targetParts) {
    // SAFETY: every target part name came out of partTexts' key set, and
    // every target part was given a spelling above
    const xml = partTexts.get(partPath)!;
    const spelling = spellings.get(partPath)!;
    const scan = scanPart({ partPath, spelling, context, seen });
    if (scan.edits.length === 0) {
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
      applySplices(xml, [...scan.edits, ...ensureRootNamespaces(spelling)], partPath),
    );
  }

  if (updates.size === 0) {
    return {
      docx: toUint8Array(docx),
      assigned: 0,
      deduplicated: 0,
      alreadyComplete: true,
      mintedParaIds: [],
    };
  }

  // Only the write-back path needs a raw JSZip (signature detection, writing
  // new part bytes, and repackaging). `docx` was already validated against
  // the entry-count/size caps above, so this second parse runs on a
  // buffer whose shape is already bounded.
  const zip = await JSZip.loadAsync(docx);

  if (hasDigitalSignatureParts(zip) && options.allowSignedPackageMutation !== true) {
    throw new EnsureParaIdsError({
      reason: ENSURE_PARA_IDS_REASONS.SIGNED_PACKAGE,
      message:
        "Refusing to normalize a digitally signed package because rewriting OOXML invalidates its signatures. Warn the user and pass allowSignedPackageMutation only if invalidation is acceptable.",
    });
  }

  for (const [partPath, content] of updates) {
    const sourceEntry = zip.file(partPath);
    if (sourceEntry === null) {
      throw createEnsureParaIdsError(`Package part disappeared during normalization: ${partPath}`);
    }
    zip.file(partPath, content, {
      compression: "DEFLATE",
      compressionOptions: { level: 6 },
      date: sourceEntry.date,
    });
  }
  const output = await zip.generateAsync({
    type: "uint8array",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });

  return { docx: output, assigned, deduplicated, alreadyComplete: false, mintedParaIds };
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
    throw createEnsureParaIdsError(`Failed to normalize paragraph IDs: ${message}`, error);
  }
};
