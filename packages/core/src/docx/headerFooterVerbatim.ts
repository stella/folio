import type { HeaderFooter } from "../types/document";

const headerFooterSerializationFingerprint = (hf: HeaderFooter): string =>
  JSON.stringify({
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
  return ext.verbatimFingerprint === headerFooterSerializationFingerprint(hf);
};

export const assignHeaderFooterVerbatimXml = (hf: HeaderFooter, xml: string): void => {
  const ext = hf;
  ext.verbatimXml = xml;
  ext.verbatimFingerprint = headerFooterSerializationFingerprint(hf);
};

export const refreshHeaderFooterVerbatimFingerprint = (hf: HeaderFooter): void => {
  const ext = hf;
  if (!ext.verbatimXml) {
    return;
  }
  ext.verbatimFingerprint = headerFooterSerializationFingerprint(hf);
};

export const clearHeaderFooterVerbatimXml = (hf: HeaderFooter): void => {
  const ext = hf;
  delete ext.verbatimXml;
  delete ext.verbatimFingerprint;
};
