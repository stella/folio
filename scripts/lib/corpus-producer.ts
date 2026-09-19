/**
 * Which tool wrote a corpus file.
 *
 * A signature that only ever fires on packages one producer writes names a
 * quirk of that producer; a signature spread across Word, LibreOffice and a
 * generator library names something folio gets wrong about the format. The
 * census reports producers per signature so the difference is visible without
 * reopening a single file.
 *
 * The evidence is the extended-properties part's `Application` and `AppVersion`
 * — the only generator hint ECMA-376 defines — plus two structural tells that
 * survive when that part is absent: the main part's name (Word Online writes
 * `word/document2.xml`) and the namespace profile of the document root. Nothing
 * else in `docProps` is read: the core properties carry author names, and the
 * gate has no business looking at them.
 *
 * Producers are reported as a bounded family plus a major version, because a
 * census keyed on the raw string would carry one key per LibreOffice build hash.
 */

import JSZip from "jszip";

export const PRODUCER_FAMILIES = {
  word: "word",
  wordMac: "word-mac",
  wordOnline: "word-online",
  outlook: "outlook",
  libreoffice: "libreoffice",
  openoffice: "openoffice",
  onlyoffice: "onlyoffice",
  wps: "wps-office",
  googleDocs: "google-docs",
  pages: "pages",
  apachePoi: "apache-poi",
  aspose: "aspose",
  docx4j: "docx4j",
  pandoc: "pandoc",
  abiword: "abiword",
  calligra: "calligra",
  /** The part is present and names an application the table does not know. */
  other: "other",
  /** The part is present but carries no `Application`. */
  unnamed: "unnamed",
  /** No extended-properties part at all: the mark of a library that built the package from nothing. */
  noExtendedProperties: "no-extended-properties",
  /** The container could not be read far enough to tell. */
  unknown: "unknown",
} as const;

export type ProducerFamily = (typeof PRODUCER_FAMILIES)[keyof typeof PRODUCER_FAMILIES];

export type CorpusProducer = {
  family: ProducerFamily;
  /** `family/major`, or `family` when no version is stated: the census key. */
  label: string;
};

const EXTENDED_PROPERTIES_PART = "docProps/app.xml";
const APPLICATION_RE = /<(?:[\w.-]+:)?Application>([^<]*)<\/(?:[\w.-]+:)?Application>/u;
const APP_VERSION_RE = /<(?:[\w.-]+:)?AppVersion>([^<]*)<\/(?:[\w.-]+:)?AppVersion>/u;
/** The leading integer of `16.0000`, `7.0.6.2$Linux_X86_64 …` or `11.1.0.11664_…`. */
const MAJOR_VERSION_RE = /(\d+)/u;

/**
 * Ordered because the tests overlap: "Microsoft Macintosh Word" also matches
 * the Word pattern, and every LibreOffice string also contains "Office".
 */
const FAMILY_PATTERNS: ReadonlyArray<{ family: ProducerFamily; pattern: RegExp }> = [
  { family: PRODUCER_FAMILIES.wordMac, pattern: /macintosh\s+word/u },
  { family: PRODUCER_FAMILIES.outlook, pattern: /outlook/u },
  { family: PRODUCER_FAMILIES.libreoffice, pattern: /libreoffice/u },
  { family: PRODUCER_FAMILIES.openoffice, pattern: /openoffice|staroffice/u },
  { family: PRODUCER_FAMILIES.onlyoffice, pattern: /onlyoffice|ascensio/u },
  { family: PRODUCER_FAMILIES.wps, pattern: /wps\s*office|kingsoft/u },
  { family: PRODUCER_FAMILIES.googleDocs, pattern: /google/u },
  { family: PRODUCER_FAMILIES.pages, pattern: /^pages\b|\bapple\b/u },
  { family: PRODUCER_FAMILIES.apachePoi, pattern: /poi|xwpf/u },
  { family: PRODUCER_FAMILIES.aspose, pattern: /aspose|groupdocs/u },
  { family: PRODUCER_FAMILIES.docx4j, pattern: /docx4j|plutext/u },
  { family: PRODUCER_FAMILIES.pandoc, pattern: /pandoc/u },
  { family: PRODUCER_FAMILIES.abiword, pattern: /abiword/u },
  { family: PRODUCER_FAMILIES.calligra, pattern: /calligra|koffice/u },
  { family: PRODUCER_FAMILIES.word, pattern: /word/u },
];

/**
 * Word Online's main part is `word/document2.xml`.
 *
 * It is the only producer that reliably names the part something other than the
 * conventional path, and it writes the same `Application` as desktop Word, so
 * the part name is the only way to tell them apart.
 */
const WORD_ONLINE_DOCUMENT_PART = "word/document2.xml";

export const producerFamilyOf = (application: string): ProducerFamily => {
  const normalized = application.trim().toLowerCase();
  if (normalized.length === 0) {
    return PRODUCER_FAMILIES.unnamed;
  }
  for (const { family, pattern } of FAMILY_PATTERNS) {
    if (pattern.test(normalized)) {
      return family;
    }
  }
  return PRODUCER_FAMILIES.other;
};

/** Families that describe the absence of evidence: a version would be a fiction. */
const VERSIONLESS_FAMILIES: ReadonlySet<ProducerFamily> = new Set([
  PRODUCER_FAMILIES.unnamed,
  PRODUCER_FAMILIES.noExtendedProperties,
  PRODUCER_FAMILIES.unknown,
]);

export const producerLabel = (family: ProducerFamily, version: string): string => {
  if (VERSIONLESS_FAMILIES.has(family)) {
    return family;
  }
  const major = MAJOR_VERSION_RE.exec(version)?.[1];
  // `00.0000` and `0.0000` are the same producer, so the label is the number.
  return major === undefined ? family : `${family}/${Number(major)}`;
};

/**
 * `AppVersion` first, then the version embedded in `Application`.
 *
 * Word states `16.0000` in `AppVersion` and nothing in `Application`;
 * LibreOffice and WPS do the reverse, and LibreOffice's `AppVersion` when it
 * writes one is Word's compatibility level rather than its own version.
 */
const WORD_FAMILIES: ReadonlySet<ProducerFamily> = new Set([
  PRODUCER_FAMILIES.word,
  PRODUCER_FAMILIES.wordMac,
  PRODUCER_FAMILIES.wordOnline,
  PRODUCER_FAMILIES.outlook,
]);

const versionFor = (family: ProducerFamily, application: string, appVersion: string): string => {
  if (WORD_FAMILIES.has(family)) {
    return appVersion;
  }
  const embedded = MAJOR_VERSION_RE.exec(application)?.[1];
  return embedded ?? appVersion;
};

export type ReadProducerOptions = {
  bytes: Uint8Array;
  /** The main part as the package relationship names it, from the classifier. */
  documentPart: string;
};

export const readCorpusProducer = async ({
  bytes,
  documentPart,
}: ReadProducerOptions): Promise<CorpusProducer> => {
  const archive = await JSZip.loadAsync(bytes).catch(() => null);
  if (archive === null) {
    return { family: PRODUCER_FAMILIES.unknown, label: PRODUCER_FAMILIES.unknown };
  }
  const part = archive.file(EXTENDED_PROPERTIES_PART);
  if (part === null) {
    return {
      family: PRODUCER_FAMILIES.noExtendedProperties,
      label: PRODUCER_FAMILIES.noExtendedProperties,
    };
  }
  const xml = await part.async("string").catch(() => null);
  if (xml === null) {
    return { family: PRODUCER_FAMILIES.unknown, label: PRODUCER_FAMILIES.unknown };
  }

  const application = APPLICATION_RE.exec(xml)?.[1] ?? "";
  const appVersion = APP_VERSION_RE.exec(xml)?.[1] ?? "";
  const named = producerFamilyOf(application);
  const family =
    named === PRODUCER_FAMILIES.word && documentPart === WORD_ONLINE_DOCUMENT_PART
      ? PRODUCER_FAMILIES.wordOnline
      : named;
  return { family, label: producerLabel(family, versionFor(family, application, appVersion)) };
};
