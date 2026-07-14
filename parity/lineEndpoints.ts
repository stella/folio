import { compareGeoms, mergeVisualRows } from "./compare";
import { normalizeLineText } from "./textNorm";
import type { Divergence, DocGeom, LineBox, PageGeom, Region } from "./types";

export const LINE_ENDPOINT_MANIFEST_SCHEMA = "folio.word-line-endpoints";
export const LINE_ENDPOINT_MANIFEST_VERSION = 1;

const SHA256_PATTERN = /^[a-f0-9]{64}$/u;

export type LineEndpoint = {
  /** Normalized text occupying one visual line. Its final token is the endpoint under test. */
  text: string;
  region: Region;
};

export type LineEndpointPage = {
  number: number;
  lines: LineEndpoint[];
};

export type WordLineEndpointManifest = {
  schema: typeof LINE_ENDPOINT_MANIFEST_SCHEMA;
  version: typeof LINE_ENDPOINT_MANIFEST_VERSION;
  source: {
    /** Basename only: manifests must not disclose a local absolute path. */
    fileName: string;
    sha256: string;
  };
  reference: {
    renderer: "word";
    capturedAt: string;
    wordVersion?: string;
    mutoolVersion?: string;
  };
  pages: LineEndpointPage[];
};

export type LineEndpointDivergence = Extract<
  Divergence,
  {
    kind:
      | "page-count"
      | "pagination"
      | "line-break"
      | "missing-line"
      | "extra-line"
      | "text-mismatch";
  }
>;

export type LineEndpointValidationResult = {
  matches: boolean;
  referenceLines: number;
  folioLines: number;
  divergences: LineEndpointDivergence[];
};

export class LineEndpointManifestError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "LineEndpointManifestError";
  }
}

type CreateManifestOptions = {
  reference: DocGeom;
  sourceFileName: string;
  sourceSha256: string;
  capturedAt?: string;
};

export const createWordLineEndpointManifest = ({
  reference,
  sourceFileName,
  sourceSha256,
  capturedAt = new Date().toISOString(),
}: CreateManifestOptions): WordLineEndpointManifest => {
  if (reference.source !== "word") {
    throw new LineEndpointManifestError("Line-endpoint manifests must be captured from Word.");
  }
  if (!SHA256_PATTERN.test(sourceSha256)) {
    throw new LineEndpointManifestError("The source SHA-256 must be 64 lowercase hex characters.");
  }
  const fileName = requireSourceFileName(sourceFileName);

  const wordVersion = nonEmpty(reference.meta["wordVersion"]);
  const mutoolVersion = nonEmpty(reference.meta["mutool"]);

  return {
    schema: LINE_ENDPOINT_MANIFEST_SCHEMA,
    version: LINE_ENDPOINT_MANIFEST_VERSION,
    source: { fileName, sha256: sourceSha256 },
    reference: {
      renderer: "word",
      capturedAt,
      ...(wordVersion === undefined ? {} : { wordVersion }),
      ...(mutoolVersion === undefined ? {} : { mutoolVersion }),
    },
    pages: extractLineEndpointPages(reference),
  };
};

export const extractLineEndpointPages = (geom: DocGeom): LineEndpointPage[] =>
  geom.pages.map((page) => ({
    number: page.number,
    lines: mergeVisualRows(page.lines)
      .map((line) => ({ text: normalizeLineText(line.text), region: line.region }))
      .filter(({ text }) => text.length > 0),
  }));

export const compareLineEndpoints = (
  manifest: WordLineEndpointManifest,
  folio: DocGeom,
): LineEndpointValidationResult => {
  const referenceGeom = endpointPagesToGeom({
    pages: manifest.pages,
    source: "word",
    file: manifest.source.fileName,
  });
  const folioPages = extractLineEndpointPages(folio);
  const folioGeom = endpointPagesToGeom({ pages: folioPages, source: "folio", file: folio.file });

  // The shared comparison engine owns sequence alignment, pagination matching,
  // and one-to-many line-break reconciliation. Canonical geometry deliberately
  // removes x/y/width differences so this validator reports endpoints only.
  const result = compareGeoms(referenceGeom, folioGeom);
  const divergences = result.divergences.filter(isLineEndpointDivergence);

  return {
    matches: divergences.length === 0,
    referenceLines: countLines(manifest.pages),
    folioLines: countLines(folioPages),
    divergences,
  };
};

type EndpointPagesToGeomOptions = {
  pages: LineEndpointPage[];
  source: "word" | "folio";
  file: string;
};

const endpointPagesToGeom = ({ pages, source, file }: EndpointPagesToGeomOptions): DocGeom => ({
  source,
  file,
  meta: { canonicalizedFor: "line-endpoints" },
  pages: pages.map(endpointPageToGeom),
});

const endpointPageToGeom = (page: LineEndpointPage): PageGeom => ({
  number: page.number,
  widthPt: 612,
  heightPt: 792,
  lines: page.lines.map(endpointToLineBox),
});

const endpointToLineBox = ({ text, region }: LineEndpoint, index: number): LineBox => ({
  text,
  normText: normalizeLineText(text),
  xPt: 0,
  yPt: index * 12,
  baselinePt: index * 12 + 9,
  widthPt: normalizeLineText(text).length,
  heightPt: 10,
  region,
});

const isLineEndpointDivergence = (divergence: Divergence): divergence is LineEndpointDivergence => {
  switch (divergence.kind) {
    case "page-count":
    case "pagination":
    case "line-break":
    case "missing-line":
    case "extra-line":
    case "text-mismatch":
      return true;
    case "x-drift":
    case "y-drift":
    case "width-drift":
      return false;
  }
};

const countLines = (pages: LineEndpointPage[]): number =>
  pages.reduce((total, page) => total + page.lines.length, 0);

const nonEmpty = (value: string | undefined): string | undefined =>
  value === undefined || value.length === 0 ? undefined : value;

export const parseLineEndpointManifest = (value: unknown): WordLineEndpointManifest => {
  const root = requireRecord(value, "manifest");
  if (root["schema"] !== LINE_ENDPOINT_MANIFEST_SCHEMA) {
    throw new LineEndpointManifestError("Unsupported line-endpoint manifest schema.");
  }
  if (root["version"] !== LINE_ENDPOINT_MANIFEST_VERSION) {
    throw new LineEndpointManifestError(
      `Unsupported line-endpoint manifest version: ${String(root["version"])}.`,
    );
  }

  const source = requireRecord(root["source"], "source");
  const fileName = requireSourceFileName(requireString(source["fileName"], "source.fileName"));
  const sha256 = requireString(source["sha256"], "source.sha256");
  if (!SHA256_PATTERN.test(sha256)) {
    throw new LineEndpointManifestError("source.sha256 must be 64 lowercase hex characters.");
  }

  const reference = requireRecord(root["reference"], "reference");
  if (reference["renderer"] !== "word") {
    throw new LineEndpointManifestError("reference.renderer must be word.");
  }
  const capturedAt = requireString(reference["capturedAt"], "reference.capturedAt");
  if (Number.isNaN(Date.parse(capturedAt))) {
    throw new LineEndpointManifestError("reference.capturedAt must be an ISO timestamp.");
  }
  const wordVersion = optionalString(reference["wordVersion"], "reference.wordVersion");
  const mutoolVersion = optionalString(reference["mutoolVersion"], "reference.mutoolVersion");

  if (!Array.isArray(root["pages"])) {
    throw new LineEndpointManifestError("pages must be an array.");
  }
  const pages = root["pages"].map(parsePage);
  for (let index = 1; index < pages.length; index += 1) {
    const previous = pages[index - 1];
    const current = pages[index];
    if (previous !== undefined && current !== undefined && current.number <= previous.number) {
      throw new LineEndpointManifestError("Page numbers must be strictly increasing.");
    }
  }

  return {
    schema: LINE_ENDPOINT_MANIFEST_SCHEMA,
    version: LINE_ENDPOINT_MANIFEST_VERSION,
    source: { fileName, sha256 },
    reference: {
      renderer: "word",
      capturedAt,
      ...(wordVersion === undefined ? {} : { wordVersion }),
      ...(mutoolVersion === undefined ? {} : { mutoolVersion }),
    },
    pages,
  };
};

const parsePage = (value: unknown, pageIndex: number): LineEndpointPage => {
  const page = requireRecord(value, `pages[${pageIndex}]`);
  const number = page["number"];
  if (typeof number !== "number" || !Number.isInteger(number) || number <= 0) {
    throw new LineEndpointManifestError(`pages[${pageIndex}].number must be a positive integer.`);
  }
  if (!Array.isArray(page["lines"])) {
    throw new LineEndpointManifestError(`pages[${pageIndex}].lines must be an array.`);
  }
  return { number, lines: page["lines"].map((line, index) => parseLine(line, pageIndex, index)) };
};

const parseLine = (value: unknown, pageIndex: number, lineIndex: number): LineEndpoint => {
  const prefix = `pages[${pageIndex}].lines[${lineIndex}]`;
  const line = requireRecord(value, prefix);
  const text = requireString(line["text"], `${prefix}.text`);
  const region = line["region"];
  if (typeof region !== "string" || !isRegion(region)) {
    throw new LineEndpointManifestError(`${prefix}.region is invalid.`);
  }
  return { text, region };
};

const isRegion = (value: string): value is Region =>
  value === "body" ||
  value === "header" ||
  value === "footer" ||
  value === "footnote" ||
  value === "unknown";

const requireRecord = (value: unknown, field: string): Record<string, unknown> => {
  if (value === null || typeof value !== "object" || Array.isArray(value)) {
    throw new LineEndpointManifestError(`${field} must be an object.`);
  }
  return Object.fromEntries(Object.entries(value));
};

const requireString = (value: unknown, field: string): string => {
  if (typeof value !== "string" || value.length === 0) {
    throw new LineEndpointManifestError(`${field} must be a non-empty string.`);
  }
  return value;
};

const optionalString = (value: unknown, field: string): string | undefined => {
  if (value === undefined) return undefined;
  return requireString(value, field);
};

const requireSourceFileName = (value: string): string => {
  if (value.length === 0 || value.includes("/") || value.includes("\\")) {
    throw new LineEndpointManifestError("source.fileName must be a basename without directories.");
  }
  return value;
};

export const readLineEndpointManifest = async (
  manifestPath: string,
): Promise<WordLineEndpointManifest> => {
  let parsed: unknown;
  try {
    parsed = JSON.parse(await Bun.file(manifestPath).text());
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    throw new LineEndpointManifestError(`Cannot read line-endpoint manifest: ${message}`);
  }
  return parseLineEndpointManifest(parsed);
};
