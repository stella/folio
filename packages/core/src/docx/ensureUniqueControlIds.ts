/**
 * Headless `w:sdtPr/w:id` normalization for a `.docx` buffer.
 *
 * A content control names itself by `w:id`, and the package model holds
 * every control by that id. A package can repeat one: a producer that copied
 * a control, or merged two documents, wrote the same id twice. The first
 * control in document order keeps its id; each later control carrying it gets
 * an id the package does not use. Every control is kept, and so is everything
 * else in the part.
 *
 * Alternate content (`mc:Fallback`) restates the controls of its `mc:Choice`
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

type ControlId = {
  /** The `w:val` attribute value's span in the part. */
  start: number;
  end: number;
  value: number;
  inFallback: boolean;
};

const isWordprocessing = (element: XmlElement, localName: string): boolean =>
  getLocalName(element.name ?? "") === localName &&
  WORDPROCESSINGML_NAMESPACE_URIS.has(getNamespaceUri(element) ?? "");

/** Every control id a part carries, in document order. */
const controlIdsIn = (xml: string): ControlId[] | undefined => {
  const out: ControlId[] = [];
  const fallbacks = new WeakSet<XmlElement>();
  const scanned = scanStreamingXmlElements(
    xml,
    ({ element, attributeValueSpans, parent }) => {
      const inFallback =
        (parent !== undefined && fallbacks.has(parent)) ||
        (getLocalName(element.name ?? "") === "Fallback" &&
          getNamespaceUri(element) === MC_NAMESPACE_URI);
      if (inFallback) fallbacks.add(element);
      if (
        parent === undefined ||
        !isWordprocessing(element, "id") ||
        !isWordprocessing(parent, "sdtPr")
      )
        return;
      for (const [name, span] of attributeValueSpans) {
        if (
          getLocalName(name) !== "val" ||
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

/** Part texts that renumber repeated ids, and how many they renumber. */
export type PackageControlIdUpdates = {
  /** The rewritten text of each part that repeats an id, by path. */
  updates: Map<string, string>;
  reminted: number;
};

/** The story parts of `paths`, the body first, then the rest in a fixed order. */
const storyPartsInOrder = (paths: Iterable<string>): string[] =>
  [...paths]
    .filter((path) => STORY_PART_PATTERN.test(path))
    .sort((a, b) => {
      const body = (path: string) =>
        path.toLowerCase() === "word/document.xml" ? 0 : 1;
      return body(a) - body(b) || a.localeCompare(b);
    });

/**
 * The part texts that give each content control after the first carrying an
 * id an id of its own; `undefined` when a part does not scan, since it may
 * hold ids a new one would repeat. The body's controls win the
 * first-occurrence rule.
 */
export const uniqueControlIdUpdates = (
  partTexts: ReadonlyMap<string, string>,
): PackageControlIdUpdates | undefined => {
  const ids = new Map<string, ControlId[]>();
  const taken = new Set<number>();
  for (const path of storyPartsInOrder(partTexts.keys())) {
    // SAFETY: the path is a key of `partTexts`
    const xml = partTexts.get(path)!;
    if (!xml.includes("sdtPr")) continue;
    const found = controlIdsIn(xml);
    if (found === undefined) return undefined;
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
  const updates = new Map<string, string>();
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
    // SAFETY: every path in `ids` is a key of `partTexts`
    const patched = spliceXml(partTexts.get(path)!, edits);
    if (patched === null) continue;
    updates.set(path, patched);
    reminted += edits.length;
  }
  return { updates, reminted };
};

export type EnsureUniqueControlIdsResult = {
  /** The normalized `.docx`; the input bytes when nothing repeated. */
  docx: Uint8Array;
  /** Controls given a new id. */
  reminted: number;
};

/**
 * Give each content control after the first carrying an id an id of its own.
 * A signed package, or a part that does not scan, is left as it is.
 */
export const ensureUniqueControlIds = async (
  docx: Uint8Array,
): Promise<EnsureUniqueControlIdsResult> => {
  const zip = await ZipArchive.read(docx);
  const entries = zip.entries();
  if (
    entries.some(({ path }) =>
      path.toLowerCase().startsWith(DIGITAL_SIGNATURE_PART_PREFIX),
    )
  )
    return { docx, reminted: 0 };
  const texts = new Map<string, string>();
  for (const path of storyPartsInOrder(entries.map((entry) => entry.path))) {
    // oxlint-disable-next-line no-await-in-loop -- parts are read one at a time in a fixed order
    const xml = await zip.entry(path)?.text();
    if (xml !== undefined) texts.set(path, xml);
  }
  const found = uniqueControlIdUpdates(texts);
  if (found === undefined || found.reminted === 0) return { docx, reminted: 0 };
  for (const [path, patched] of found.updates) {
    const source = zip.entry(path);
    zip.set(path, patched, source ? { modified: source.modified } : {});
  }
  return { docx: await zip.generate({ level: 6 }), reminted: found.reminted };
};
