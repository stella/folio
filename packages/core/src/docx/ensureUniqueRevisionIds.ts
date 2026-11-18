/**
 * Headless revision `w:id` normalization for a `.docx` buffer.
 *
 * Each tracked change (`w:ins`, `w:del`, a property change, a table cell
 * change) carries a `w:id`, and the package model holds every revision by it.
 * A package can repeat one: a producer that numbered per part, copied a
 * change, or wrote one id across the records of one edit. Each record is
 * still a change of its own. The first record in
 * document order keeps its id; each later record carrying it gets an id the
 * package does not use. Every record is kept, and so is everything else in
 * the part. Range markers (bookmarks, comments, move and custom XML ranges)
 * pair their start and end by id and are left alone.
 *
 * Alternate content (`mc:Fallback`) restates the records of its `mc:Choice`
 * under the same ids for older readers, and the parser reads one of the two,
 * so a fallback's ids are reserved but never counted as a repeat.
 */

import { scanStreamingXmlElements } from "@stll/docx-core/xml";
import { ZipArchive } from "@stll/docx-core/zip";

import { spliceXml } from "./selectiveXmlPatch";
import {
  getLocalName,
  getNamespaceUri,
  resolveAttributeNamespaceUri,
  WORDPROCESSINGML_NAMESPACE_URIS,
  type XmlElement,
} from "./xmlParser";

const MC_NAMESPACE_URI =
  "http://schemas.openxmlformats.org/markup-compatibility/2006";
const DIGITAL_SIGNATURE_PART_PREFIX = "_xmlsignatures/";
/**
 * The parts the loader can read records from: the body, headers, footers,
 * notes and comments wherever their relationships place them under `word/`,
 * and the styles and numbering their property changes live in.
 */
const STORY_PART_PATTERN = /^word\/.+\.xml$/iu;

/** The elements that are revision records, each with an id of its own. */
const REVISION_RECORDS: ReadonlySet<string> = new Set([
  "ins",
  "del",
  "moveFrom",
  "moveTo",
  "rPrChange",
  "pPrChange",
  "sectPrChange",
  "tblPrChange",
  "tblPrExChange",
  "tblGridChange",
  "trPrChange",
  "tcPrChange",
  "cellIns",
  "cellDel",
  "cellMerge",
  "numberingChange",
]);

type RevisionId = {
  /** The `w:id` attribute value's span in the part. */
  start: number;
  end: number;
  value: number;
  inFallback: boolean;
};

/** Every revision record id a part carries, in document order. */
const revisionIdsIn = (xml: string): RevisionId[] | undefined => {
  const out: RevisionId[] = [];
  const fallbacks = new WeakSet<XmlElement>();
  const scanned = scanStreamingXmlElements(
    xml,
    ({ element, attributeValueSpans, parent }) => {
      const localName = getLocalName(element.name ?? "");
      const inFallback =
        (parent !== undefined && fallbacks.has(parent)) ||
        (localName === "Fallback" &&
          getNamespaceUri(element) === MC_NAMESPACE_URI);
      if (inFallback) fallbacks.add(element);
      if (
        !REVISION_RECORDS.has(localName) ||
        !WORDPROCESSINGML_NAMESPACE_URIS.has(getNamespaceUri(element) ?? "")
      )
        return;
      for (const [name, span] of attributeValueSpans) {
        if (
          getLocalName(name) !== "id" ||
          !WORDPROCESSINGML_NAMESPACE_URIS.has(
            resolveAttributeNamespaceUri(element, name) ?? "",
          )
        )
          continue;
        const value = Number(element.attributes?.[name]);
        if (Number.isInteger(value))
          out.push({ start: span.start, end: span.end, value, inFallback });
      }
    },
  );
  return scanned.status === "scanned" ? out : undefined;
};

export type EnsureUniqueRevisionIdsResult = {
  /** The normalized `.docx`; the input bytes when nothing repeated. */
  docx: Uint8Array;
  /** Revision records given a new id. */
  reminted: number;
};

/**
 * Give each revision record after the first carrying an id an id of its own.
 * A signed package, or a part that does not scan, is left as it is.
 */
export const ensureUniqueRevisionIds = async (
  docx: Uint8Array,
): Promise<EnsureUniqueRevisionIdsResult> => {
  const zip = await ZipArchive.read(docx);
  const entries = zip.entries();
  if (
    entries.some(({ path }) =>
      path.toLowerCase().startsWith(DIGITAL_SIGNATURE_PART_PREFIX),
    )
  )
    return { docx, reminted: 0 };
  // The body first, so its records win the first-occurrence rule, then the
  // other story parts in a fixed order.
  const parts = entries
    .map(({ path }) => path)
    .filter((path) => STORY_PART_PATTERN.test(path))
    .sort((a, b) => {
      const body = (path: string) =>
        path.toLowerCase() === "word/document.xml" ? 0 : 1;
      return body(a) - body(b) || a.localeCompare(b);
    });
  const texts = new Map<string, string>();
  const ids = new Map<string, RevisionId[]>();
  const taken = new Set<number>();
  for (const path of parts) {
    // oxlint-disable-next-line no-await-in-loop -- parts are read one at a time in a fixed order
    const xml = await zip.entry(path)?.text();
    if (xml === undefined) continue;
    const found = revisionIdsIn(xml);
    // A part the scanner cannot read may hold ids a new one would repeat.
    if (found === undefined) return { docx, reminted: 0 };
    if (found.length === 0) continue;
    texts.set(path, xml);
    ids.set(path, found);
    for (const { value } of found) taken.add(value);
  }
  let next = 1;
  const fresh = (): number => {
    while (taken.has(next)) next += 1;
    taken.add(next);
    return next;
  };
  const seen = new Set<number>();
  let reminted = 0;
  for (const [path, found] of ids) {
    const edits: { start: number; end: number; newXml: string }[] = [];
    for (const { start, end, value, inFallback } of found) {
      if (inFallback) continue;
      if (!seen.has(value)) {
        seen.add(value);
        continue;
      }
      edits.push({ start, end, newXml: String(fresh()) });
    }
    if (edits.length === 0) continue;
    // SAFETY: every path in `ids` has its text in `texts`
    const patched = spliceXml(texts.get(path)!, edits);
    if (patched === null) continue;
    const source = zip.entry(path);
    zip.set(path, patched, source ? { modified: source.modified } : {});
    reminted += edits.length;
  }
  if (reminted === 0) return { docx, reminted };
  return { docx: await zip.generate({ level: 6 }), reminted };
};
