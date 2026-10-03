import type { BlockContent, HeaderFooter } from "../types/document";
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
  { fingerprint: string; content: readonly BlockContent[] }
>();

const captureContentBaseline = (hf: HeaderFooter): void => {
  const fingerprint = hf.verbatimFingerprint;
  if (fingerprint === undefined) return;
  const handle = {};
  contentBaselines.set(handle, { fingerprint, content: structuredClone(hf.content) });
  Object.defineProperty(hf, BASELINE_HANDLE, {
    value: handle,
    enumerable: true,
    configurable: true,
  });
};

export const getHeaderFooterBaselineContent = (
  hf: HeaderFooter,
): readonly BlockContent[] | undefined => {
  if (!(BASELINE_HANDLE in hf)) return;
  const handle = hf[BASELINE_HANDLE];
  if (typeof handle !== "object" || handle === null) return;
  const baseline = contentBaselines.get(handle);
  return baseline?.fingerprint === hf.verbatimFingerprint ? baseline?.content : undefined;
};

export const canReplayHeaderFooterBlocks = (hf: HeaderFooter): boolean => {
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
