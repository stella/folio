import type { BlockContent, Document, HeaderFooter } from "../types/document";
import { Result } from "better-result";
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
const BASELINE_HANDLE = Symbol("header-footer-source-baseline");
const contentBaselines = new WeakMap<
  object,
  {
    fingerprint: string;
    xml: string | undefined;
    contentFingerprint: string;
    content: readonly BlockContent[];
  }
>();

const captureContentBaseline = (hf: HeaderFooter): void => {
  const fingerprint = hf.verbatimFingerprint;
  if (fingerprint === undefined) return;
  const handle = {};
  const content = structuredClone(hf.content);
  contentBaselines.set(handle, {
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
