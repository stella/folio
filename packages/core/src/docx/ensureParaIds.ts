/**
 * Headless `w14:paraId` normalization for a `.docx` buffer.
 *
 * Word 2010+ stamps every paragraph with a `w14:paraId`, but Google Docs
 * exports, LibreOffice, python-docx, and docx4j generally do not. Paragraphs
 * without one are invisible to everything that anchors on paraId (block ids,
 * comment threads, AI-edit targeting) and fall back to positional `seq-NNNN`
 * ids that renumber after structural edits. Hosts call {@link ensureParaIds}
 * once at ingest so every stored version has full id coverage before any
 * snapshot is taken.
 *
 * The pass patches part XML in place (string splices, no model round-trip),
 * so documents carrying features folio's parser does not model come back
 * with those features byte-identical. Contract:
 *
 * - Parts covered: `word/document.xml`, `word/header*.xml`, `word/footer*.xml`,
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
 * - Idempotent: a document that already has full coverage is returned as the
 *   original bytes, untouched (`alreadyComplete: true`).
 */
import JSZip from "jszip";

import { deterministicHexId } from "../utils/hexId";
import { isXmlNameBoundary } from "./selectiveXmlPatch";

export class EnsureParaIdsError extends Error {
  override name = "EnsureParaIdsError";
}

export type EnsureParaIdsResult = {
  /** The normalized `.docx`; the input bytes verbatim when `alreadyComplete`. */
  docx: Uint8Array;
  /** Paragraphs that received a paraId (missing or all-zero before). */
  assigned: number;
  /** Duplicate paraIds reassigned (the first occurrence keeps the id). */
  deduplicated: number;
  /** True when the input already had full, unique coverage. */
  alreadyComplete: boolean;
};

const W14_NAMESPACE_URI = "http://schemas.microsoft.com/office/word/2010/wordml";
const MC_NAMESPACE_URI = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const MC_IGNORABLE_W14 = "w14";
/** Word reads an all-zero `w14:paraId` as "no id assigned". */
const RESERVED_ZERO_ID_PATTERN = /^0*$/u;

const DOCUMENT_PART = "word/document.xml";
const TARGET_PART_PATTERNS = [
  /^word\/document\.xml$/u,
  /^word\/header\d*\.xml$/u,
  /^word\/footer\d*\.xml$/u,
  /^word\/footnotes\.xml$/u,
  /^word\/endnotes\.xml$/u,
];

const PARAGRAPH_OPEN = "<w:p";
const FALLBACK_OPEN = "<mc:Fallback";
const FALLBACK_CLOSE = "</mc:Fallback>";

/**
 * The parser accepts the `w:` fallback prefix for paraId, and comment parts
 * link replies via `w15:paraId` — all three count toward uniqueness.
 */
const ANY_PARA_ID_PATTERN = /\bw(?:14|15)?:paraId="(?<id>[^"]*)"/gu;
const OPEN_TAG_PARA_ID_PATTERN = /\sw(?:14)?:paraId="(?<id>[^"]*)"/u;
const OPEN_TAG_TEXT_ID_PATTERN = /\sw(?:14)?:textId="/u;
const XMLNS_W14_PATTERN = /\sxmlns:w14="/u;
const XMLNS_MC_PATTERN = /\sxmlns:mc="/u;
const MC_IGNORABLE_PATTERN = /\smc:Ignorable="(?<value>[^"]*)"/u;

type SpliceEdit = { start: number; end: number; text: string };

const applySplices = (xml: string, edits: SpliceEdit[]): string => {
  const ordered = [...edits].sort((a, b) => b.start - a.start);
  let result = xml;
  for (const { start, end, text } of ordered) {
    result = result.slice(0, start) + text + result.slice(end);
  }
  return result;
};

const collectExistingParaIds = (xml: string, into: Set<string>): void => {
  for (const match of xml.matchAll(ANY_PARA_ID_PATTERN)) {
    // SAFETY: named group `id` always present when the pattern matches
    const id = match.groups!["id"]!;
    if (id.length > 0) {
      into.add(id.toUpperCase());
    }
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
  assigned: number;
  deduplicated: number;
};

/**
 * One forward scan over a part: every `<w:p>` open tag outside `mc:Fallback`
 * either keeps its paraId (valid, first occurrence) or gets a splice edit
 * minting / replacing one. `seen` spans parts so the first-occurrence rule is
 * document-wide across the scan order.
 */
const scanPart = (
  xml: string,
  partPath: string,
  context: MintContext,
  seen: Set<string>,
): PartScanResult => {
  const edits: SpliceEdit[] = [];
  let assigned = 0;
  let deduplicated = 0;
  let ordinal = 0;
  let fallbackDepth = 0;
  let pos = 0;

  while (pos < xml.length) {
    const tagStart = xml.indexOf("<", pos);
    if (tagStart === -1) {
      break;
    }

    if (
      xml.startsWith(FALLBACK_OPEN, tagStart) &&
      isXmlNameBoundary(xml[tagStart + FALLBACK_OPEN.length])
    ) {
      const tagEnd = xml.indexOf(">", tagStart);
      if (tagEnd === -1) {
        throw new EnsureParaIdsError(`Unterminated <mc:Fallback> tag in ${partPath}`);
      }
      if (xml[tagEnd - 1] !== "/") {
        fallbackDepth += 1;
      }
      pos = tagEnd + 1;
      continue;
    }

    if (xml.startsWith(FALLBACK_CLOSE, tagStart)) {
      fallbackDepth = Math.max(0, fallbackDepth - 1);
      pos = tagStart + FALLBACK_CLOSE.length;
      continue;
    }

    if (
      !xml.startsWith(PARAGRAPH_OPEN, tagStart) ||
      !isXmlNameBoundary(xml[tagStart + PARAGRAPH_OPEN.length])
    ) {
      pos = tagStart + 1;
      continue;
    }

    const tagEnd = xml.indexOf(">", tagStart);
    if (tagEnd === -1) {
      throw new EnsureParaIdsError(`Unterminated <w:p> tag in ${partPath}`);
    }
    pos = tagEnd + 1;
    if (fallbackDepth > 0) {
      continue;
    }
    ordinal += 1;

    const openTag = xml.slice(tagStart, tagEnd + 1);
    const paraIdMatch = OPEN_TAG_PARA_ID_PATTERN.exec(openTag);
    if (!paraIdMatch) {
      const id = mintParaId(context, partPath, ordinal);
      const textId = OPEN_TAG_TEXT_ID_PATTERN.test(openTag) ? "" : ` w14:textId="${id}"`;
      const insertAt = tagStart + PARAGRAPH_OPEN.length;
      edits.push({ start: insertAt, end: insertAt, text: ` w14:paraId="${id}"${textId}` });
      assigned += 1;
      seen.add(id);
      continue;
    }

    // SAFETY: named group `id` always present when the pattern matches
    const rawValue = paraIdMatch.groups!["id"]!;
    const value = rawValue.toUpperCase();
    const unassigned = RESERVED_ZERO_ID_PATTERN.test(value);
    if (!unassigned && !seen.has(value)) {
      seen.add(value);
      continue;
    }

    const id = mintParaId(context, partPath, ordinal);
    const valueStart = tagStart + paraIdMatch.index + paraIdMatch[0].indexOf('"') + 1;
    edits.push({ start: valueStart, end: valueStart + rawValue.length, text: id });
    if (unassigned) {
      assigned += 1;
    } else {
      deduplicated += 1;
    }
    seen.add(id);
  }

  return { edits, assigned, deduplicated };
};

/**
 * Splice edits ensuring the part's root element declares `xmlns:w14` /
 * `xmlns:mc` and lists `w14` in `mc:Ignorable`, so consumers that predate the
 * 2010 extensions skip the new attributes instead of rejecting the part.
 */
const ensureRootNamespaces = (xml: string, partPath: string): SpliceEdit[] => {
  let pos = 0;
  let rootStart = -1;
  while (pos < xml.length) {
    const lt = xml.indexOf("<", pos);
    if (lt === -1) {
      break;
    }
    const next = xml[lt + 1];
    if (next === "?" || next === "!") {
      const skipTo = xml.startsWith("<!--", lt) ? xml.indexOf("-->", lt) : xml.indexOf(">", lt);
      if (skipTo === -1) {
        throw new EnsureParaIdsError(`Unterminated prolog in ${partPath}`);
      }
      pos = skipTo + 1;
      continue;
    }
    rootStart = lt;
    break;
  }
  if (rootStart === -1) {
    throw new EnsureParaIdsError(`No root element in ${partPath}`);
  }
  const rootEnd = xml.indexOf(">", rootStart);
  if (rootEnd === -1) {
    throw new EnsureParaIdsError(`Unterminated root element in ${partPath}`);
  }
  const rootTag = xml.slice(rootStart, rootEnd + 1);

  const edits: SpliceEdit[] = [];
  const declarations: string[] = [];
  if (!XMLNS_MC_PATTERN.test(rootTag)) {
    declarations.push(` xmlns:mc="${MC_NAMESPACE_URI}"`);
  }
  if (!XMLNS_W14_PATTERN.test(rootTag)) {
    declarations.push(` xmlns:w14="${W14_NAMESPACE_URI}"`);
  }

  const ignorable = MC_IGNORABLE_PATTERN.exec(rootTag);
  if (ignorable) {
    // SAFETY: named group `value` always present when the pattern matches
    const value = ignorable.groups!["value"]!;
    const tokens = value.split(/\s+/u).filter((token) => token.length > 0);
    if (!tokens.includes(MC_IGNORABLE_W14)) {
      const valueStart = rootStart + ignorable.index + ignorable[0].indexOf('"') + 1;
      const valueEnd = valueStart + value.length;
      const text = tokens.length === 0 ? MC_IGNORABLE_W14 : ` ${MC_IGNORABLE_W14}`;
      edits.push({ start: valueEnd, end: valueEnd, text });
    }
  } else {
    declarations.push(` mc:Ignorable="${MC_IGNORABLE_W14}"`);
  }

  if (declarations.length > 0) {
    let nameEnd = rootStart + 1;
    while (nameEnd < xml.length && !isXmlNameBoundary(xml[nameEnd])) {
      nameEnd += 1;
    }
    edits.push({ start: nameEnd, end: nameEnd, text: declarations.join("") });
  }

  return edits;
};

/** OPC part names are case-insensitive; compare lowercased. */
const isTargetPart = (path: string): boolean =>
  TARGET_PART_PATTERNS.some((pattern) => pattern.test(path.toLowerCase()));

const toUint8Array = (docx: Uint8Array | ArrayBuffer): Uint8Array =>
  docx instanceof Uint8Array ? docx : new Uint8Array(docx);

/**
 * Backfill `w14:paraId` on every paragraph of a `.docx` buffer. See the
 * module doc for the exact contract. Throws {@link EnsureParaIdsError} when
 * the buffer is not a WordprocessingML package or a part is malformed.
 */
export const ensureParaIds = async (
  docx: Uint8Array | ArrayBuffer,
): Promise<EnsureParaIdsResult> => {
  const zip = await JSZip.loadAsync(docx);

  const xmlPartNames: string[] = [];
  zip.forEach((relativePath, entry) => {
    const lower = relativePath.toLowerCase();
    if (!entry.dir && lower.startsWith("word/") && lower.endsWith(".xml")) {
      xmlPartNames.push(relativePath);
    }
  });
  const documentPartName = xmlPartNames.find((name) => name.toLowerCase() === DOCUMENT_PART);
  if (documentPartName === undefined) {
    throw new EnsureParaIdsError("word/document.xml not found — not a WordprocessingML package");
  }

  const partTexts = new Map<string, string>();
  for (const name of xmlPartNames) {
    const entry = zip.file(name);
    if (entry) {
      partTexts.set(name, await entry.async("text"));
    }
  }

  const taken = new Set<string>();
  for (const text of partTexts.values()) {
    collectExistingParaIds(text, taken);
  }

  // SAFETY: documentPartName came out of partTexts' key set
  const documentXml = partTexts.get(documentPartName)!;
  const context: MintContext = { taken, docKey: deterministicHexId(documentXml) };

  // document.xml first so its paragraphs win the first-occurrence rule, then
  // the auxiliary parts in a deterministic order.
  const targetParts = [
    documentPartName,
    ...xmlPartNames
      .filter((name) => name !== documentPartName && isTargetPart(name))
      .sort((a, b) => a.localeCompare(b)),
  ];

  const seen = new Set<string>();
  const updates = new Map<string, string>();
  let assigned = 0;
  let deduplicated = 0;

  for (const partPath of targetParts) {
    // SAFETY: every target part name came out of partTexts' key set
    const xml = partTexts.get(partPath)!;
    const scan = scanPart(xml, partPath, context, seen);
    if (scan.edits.length === 0) {
      continue;
    }
    assigned += scan.assigned;
    deduplicated += scan.deduplicated;
    updates.set(
      partPath,
      applySplices(xml, [...scan.edits, ...ensureRootNamespaces(xml, partPath)]),
    );
  }

  if (updates.size === 0) {
    return { docx: toUint8Array(docx), assigned: 0, deduplicated: 0, alreadyComplete: true };
  }

  for (const [partPath, content] of updates) {
    zip.file(partPath, content, { compression: "DEFLATE", compressionOptions: { level: 6 } });
  }
  const output = await zip.generateAsync({
    type: "uint8array",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });

  return { docx: output, assigned, deduplicated, alreadyComplete: false };
};
