import type { BlockContent, Document, HeaderFooter } from "../types/document";
import { Result, panic } from "better-result";
import { canonicalJson } from "../utils/canonicalJson";

const fingerprintBaselines = new WeakMap<HeaderFooter, { fingerprint: string; value: unknown }>();

const readFingerprint = (hf: HeaderFooter): unknown => {
  const fingerprint = hf.verbatimFingerprint;
  if (!fingerprint) return undefined;
  const cached = fingerprintBaselines.get(hf);
  if (cached?.fingerprint === fingerprint) return cached.value;
  const parsed = Result.try((): unknown => JSON.parse(fingerprint));
  const value = parsed.isOk() ? parsed.value : undefined;
  fingerprintBaselines.set(hf, { fingerprint, value });
  return value;
};

// The capture handle survives object spreads without exposing baseline content in JSON.
const BASELINE_HANDLE = Symbol("package-source-baseline");
const contentBaselines = new WeakMap<
  object,
  | {
      type: "headerFooter";
      fingerprint: string;
      xml: string | undefined;
      contentFingerprint: string;
      content: readonly BlockContent[];
    }
  | {
      type: "body";
      fingerprint: { type: "pending" } | { type: "captured"; value: string };
      kind: "mainPart";
      xml: string;
      body: Document["package"]["document"];
      resourceStyles: Readonly<NonNullable<Document["package"]["styles"]>> | undefined;
      styleIds: ReadonlySet<string>;
      resourceRelationships: Document["package"]["relationships"];
      resourceMedia: ReadonlyMap<string, { data: ArrayBuffer; mimeType: string }>;
    }
  | {
      type: "body";
      kind: "noMainPart";
      resourceStyles: Readonly<NonNullable<Document["package"]["styles"]>> | undefined;
      styleIds: ReadonlySet<string>;
    }
>();

const captureContentBaseline = (hf: HeaderFooter): void => {
  const fingerprint = hf.verbatimFingerprint;
  if (fingerprint === undefined) return;
  const handle = {};
  const content = structuredClone(hf.content);
  contentBaselines.set(handle, {
    type: "headerFooter",
    fingerprint,
    xml: hf.verbatimXml,
    contentFingerprint: canonicalJson(content),
    content,
  });
  Object.defineProperty(hf, BASELINE_HANDLE, {
    value: handle,
    enumerable: true,
    configurable: true,
  });
};

type HeaderFooterSourceBaseline =
  | { type: "missing" }
  | { type: "mismatch" }
  | { type: "captured"; content: readonly BlockContent[] };

// A story journal is JSON data; its restored records can recover ownership from
// the same parsed package without trusting their fingerprint as model data.
const packageBaselines = new WeakMap<ArrayBuffer, Map<string, Map<string, object>>>();
export const captureHeaderFooterPackageBaselines = (document: Document): void => {
  if (!document.originalBuffer) return;
  const handles = new Map<string, Map<string, object>>();
  for (const parts of [document.package.headers, document.package.footers]) {
    for (const hf of parts?.values() ?? []) {
      if (
        !(BASELINE_HANDLE in hf) ||
        typeof hf[BASELINE_HANDLE] !== "object" ||
        hf[BASELINE_HANDLE] === null ||
        hf.verbatimFingerprint === undefined ||
        hf.verbatimXml === undefined
      )
        continue;
      const bySource = handles.get(hf.verbatimFingerprint) ?? new Map<string, object>();
      bySource.set(hf.verbatimXml, hf[BASELINE_HANDLE]);
      handles.set(hf.verbatimFingerprint, bySource);
    }
  }
  packageBaselines.set(document.originalBuffer, handles);
};

/**
 * Transfer body, header and footer capture handles across a trusted
 * document graph clone. `structuredClone` drops symbol-keyed properties, so a
 * clone would otherwise read an edited part as uncaptured ("missing") rather
 * than as a mismatch against its source.
 */
export const copyHeaderFooterBaselineHandles = (target: Document, source: Document): void => {
  const sourceBody = source.package.document;
  if (BASELINE_HANDLE in sourceBody)
    Object.defineProperty(target.package.document, BASELINE_HANDLE, {
      value: sourceBody[BASELINE_HANDLE],
      enumerable: true,
      configurable: true,
    });
  for (const kind of ["headers", "footers"] as const) {
    const targets = target.package[kind];
    for (const [rId, sourcePart] of source.package[kind] ?? []) {
      if (!(BASELINE_HANDLE in sourcePart)) continue;
      const targetPart = targets?.get(rId);
      if (!targetPart) continue;
      Object.defineProperty(targetPart, BASELINE_HANDLE, {
        value: sourcePart[BASELINE_HANDLE],
        enumerable: true,
        configurable: true,
      });
    }
  }
};

/** Transfer the parsed package registry only across a trusted document graph clone. */
export const copyHeaderFooterPackageBaselines = (target: Document, source: Document): void => {
  if (!target.originalBuffer || !source.originalBuffer) return;
  const baselines = packageBaselines.get(source.originalBuffer);
  if (baselines) packageBaselines.set(target.originalBuffer, baselines);
};

export const getHeaderFooterSourceBaseline = (
  hf: HeaderFooter,
  originalBuffer?: ArrayBuffer,
): HeaderFooterSourceBaseline => {
  let handle: unknown;
  if (BASELINE_HANDLE in hf) handle = hf[BASELINE_HANDLE];
  else if (originalBuffer && hf.verbatimFingerprint !== undefined && hf.verbatimXml !== undefined)
    handle = packageBaselines.get(originalBuffer)?.get(hf.verbatimFingerprint)?.get(hf.verbatimXml);
  if (handle === undefined && !(BASELINE_HANDLE in hf)) return { type: "missing" };
  if (typeof handle !== "object" || handle === null) return { type: "mismatch" };
  const baseline = contentBaselines.get(handle);
  if (
    !baseline ||
    baseline.type !== "headerFooter" ||
    baseline.fingerprint !== hf.verbatimFingerprint ||
    baseline.xml !== hf.verbatimXml ||
    baseline.contentFingerprint !== canonicalJson(baseline.content)
  )
    return { type: "mismatch" };
  return { type: "captured", content: baseline.content };
};

export const canReplayHeaderFooterBlocks = (
  hf: HeaderFooter,
  baselineContent: readonly BlockContent[],
): boolean => {
  if (hf.rawWatermarkXml !== undefined) {
    const index = hf.watermarkBlockIndex;
    if (
      index === undefined ||
      !Number.isInteger(index) ||
      index < 0 ||
      index >= baselineContent.length ||
      baselineContent.length !== hf.content.length ||
      canonicalJson(baselineContent.at(index)) !== canonicalJson(hf.content.at(index))
    )
      return false;
  }
  const baseline = readFingerprint(hf);
  if (
    baseline === null ||
    typeof baseline !== "object" ||
    !("type" in baseline) ||
    baseline.type !== hf.type
  )
    return false;
  return (
    canonicalJson({
      watermark: "watermark" in baseline ? baseline.watermark : undefined,
      watermarkBlockIndex:
        "watermarkBlockIndex" in baseline ? baseline.watermarkBlockIndex : undefined,
      rawWatermarkXml: "rawWatermarkXml" in baseline ? baseline.rawWatermarkXml : undefined,
    }) ===
    canonicalJson({
      watermark: hf.watermark,
      watermarkBlockIndex: hf.watermarkBlockIndex,
      rawWatermarkXml: hf.rawWatermarkXml,
    })
  );
};

const headerFooterSerializationFingerprint = (hf: HeaderFooter): string =>
  JSON.stringify({
    type: hf.type,
    content: hf.content,
    watermark: hf.watermark,
    watermarkBlockIndex: hf.watermarkBlockIndex,
    rawWatermarkXml: hf.rawWatermarkXml,
  });

export const getHeaderFooterVerbatimXml = (hf: HeaderFooter): string | undefined => hf.verbatimXml;

export const canReplayHeaderFooterVerbatim = (hf: HeaderFooter): boolean => {
  const ext = hf;
  if (!ext.verbatimXml || !ext.verbatimFingerprint) {
    return false;
  }
  return (
    canonicalJson(readFingerprint(hf)) ===
    canonicalJson({
      type: hf.type,
      content: hf.content,
      watermark: hf.watermark,
      watermarkBlockIndex: hf.watermarkBlockIndex,
      rawWatermarkXml: hf.rawWatermarkXml,
    })
  );
};

export const assignHeaderFooterVerbatimXml = (hf: HeaderFooter, xml: string): void => {
  const ext = hf;
  ext.verbatimXml = xml;
  ext.verbatimFingerprint = headerFooterSerializationFingerprint(hf);
  captureContentBaseline(hf);
};

export const refreshHeaderFooterVerbatimFingerprint = (hf: HeaderFooter): void => {
  const ext = hf;
  if (!ext.verbatimXml) {
    return;
  }
  ext.verbatimFingerprint = headerFooterSerializationFingerprint(hf);
  captureContentBaseline(hf);
};

export const clearHeaderFooterVerbatimXml = (hf: HeaderFooter): void => {
  const ext = hf;
  delete ext.verbatimXml;
  delete ext.verbatimFingerprint;
  Reflect.deleteProperty(hf, BASELINE_HANDLE);
};

/** Freeze the parsed plain-object graph once, retaining shared references. */
const freezeSourceStyles = (value: unknown, visited = new WeakSet<object>()): void => {
  if (typeof value !== "object" || value === null || visited.has(value)) return;
  visited.add(value);
  for (const child of Object.values(value)) freezeSourceStyles(child, visited);
  Object.freeze(value);
};

/** Capture parsed resources on the same trusted handle, including an absent main part. */
export const captureDocumentSourceBaseline = (
  document: Document,
  xml: string | undefined,
): void => {
  const handle = {};
  const resourceStyles = structuredClone(document.package.styles);
  freezeSourceStyles(resourceStyles);
  const styleIds = new Set(resourceStyles?.styles.map(({ styleId }) => styleId));
  if (xml === undefined) {
    contentBaselines.set(handle, { type: "body", kind: "noMainPart", resourceStyles, styleIds });
  } else {
    contentBaselines.set(handle, {
      type: "body",
      kind: "mainPart",
      fingerprint: { type: "pending" },
      xml,
      body: structuredClone(document.package.document),
      resourceStyles,
      styleIds,
      resourceRelationships: structuredClone(document.package.relationships),
      resourceMedia: new Map(
        [...(document.package.media ?? [])].map(([path, media]) => [
          path,
          { data: media.data, mimeType: media.mimeType },
        ]),
      ),
    });
  }
  Object.defineProperty(document.package.document, BASELINE_HANDLE, {
    value: handle,
    enumerable: true,
    configurable: true,
  });
  if (document.originalBuffer) {
    const registry = packageBaselines.get(document.originalBuffer) ?? new Map();
    registry.set("word/document.xml", new Map([["body", handle]]));
    packageBaselines.set(document.originalBuffer, registry);
  }
};

// Styles remain private in the same capture as the parsed document body.
const readDocumentSourceCapture = (document: Document) => {
  const body = document.package.document;
  const handle =
    BASELINE_HANDLE in body
      ? body[BASELINE_HANDLE]
      : document.originalBuffer &&
        packageBaselines.get(document.originalBuffer)?.get("word/document.xml")?.get("body");
  if (handle === undefined) return { type: "missing" } as const;
  if (typeof handle !== "object" || handle === null) return { type: "mismatch" } as const;
  const baseline = contentBaselines.get(handle);
  if (!baseline || baseline.type !== "body") return { type: "mismatch" } as const;
  return { type: "captured", baseline } as const;
};

/** Read the identity-stable, deeply frozen stylesheet without cloning or body comparison. */
export const getDocumentSourceStyles = (
  document: Document,
  liveStyles?: Document["package"]["styles"],
) => {
  const source = readDocumentSourceCapture(document);
  if (source.type !== "captured") return source;
  const additions =
    liveStyles?.styles.filter(({ styleId }) => !source.baseline.styleIds.has(styleId)) ?? [];
  return { type: "captured", styles: source.baseline.resourceStyles, additions } as const;
};

export const getDocumentSourceBaseline = (
  document: Document,
):
  | { type: "missing" }
  | { type: "mismatch" }
  | {
      type: "captured";
      body: Document["package"]["document"];
      xml: string;
      fingerprint: string;
      resourceStyles: Document["package"]["styles"];
      resourceRelationships: Document["package"]["relationships"];
      resourceMedia: ReadonlyMap<string, { data: ArrayBuffer; mimeType: string }>;
    } => {
  const source = readDocumentSourceCapture(document);
  if (source.type !== "captured") return source;
  const { baseline } = source;
  switch (baseline.kind) {
    case "noMainPart":
      return { type: "missing" };
    case "mainPart":
      break;
    default: {
      const unreachable: never = baseline;
      return panic(`Unexpected document source capture: ${JSON.stringify(unreachable)}`);
    }
  }
  const fingerprint = JSON.stringify(baseline.body);
  // The cloned body stays private until this first read. Capture its integrity
  // fingerprint lazily so parsing alone never pays for save-only comparisons.
  if (baseline.fingerprint.type === "pending")
    baseline.fingerprint = { type: "captured", value: fingerprint };
  if (baseline.fingerprint.value !== fingerprint) return { type: "mismatch" };
  return {
    type: "captured",
    body: baseline.body,
    xml: baseline.xml,
    fingerprint,
    resourceStyles: baseline.resourceStyles,
    resourceRelationships: baseline.resourceRelationships,
    resourceMedia: baseline.resourceMedia,
  };
};
