import type { HeaderFooter } from "../types/document";

/**
 * Folio extension on {@link HeaderFooter}: original part XML captured at parse
 * time so unedited headers/footers re-emit byte-identically on save (VML OLE
 * wrappers, smart tags, and other constructs the model cannot fully represent).
 * Cleared on first edit.
 */
export type HeaderFooterWithVerbatim = HeaderFooter & {
  verbatimXml?: string;
  /** Fingerprint of modeled fields at parse time; verbatim replay is safe only while it matches. */
  verbatimFingerprint?: string;
};

const headerFooterSerializationFingerprint = (hf: HeaderFooter): string =>
  JSON.stringify({
    content: hf.content,
    watermark: hf.watermark,
    watermarkBlockIndex: hf.watermarkBlockIndex,
    rawWatermarkXml: hf.rawWatermarkXml,
  });

export const getHeaderFooterVerbatimXml = (hf: HeaderFooter): string | undefined =>
  (hf as HeaderFooterWithVerbatim).verbatimXml;

export const canReplayHeaderFooterVerbatim = (hf: HeaderFooter): boolean => {
  const ext = hf as HeaderFooterWithVerbatim;
  if (!ext.verbatimXml || !ext.verbatimFingerprint) {
    return false;
  }
  return ext.verbatimFingerprint === headerFooterSerializationFingerprint(hf);
};

export const assignHeaderFooterVerbatimXml = (hf: HeaderFooter, xml: string): void => {
  const ext = hf as HeaderFooterWithVerbatim;
  ext.verbatimXml = xml;
  ext.verbatimFingerprint = headerFooterSerializationFingerprint(hf);
};

export const clearHeaderFooterVerbatimXml = (hf: HeaderFooter): void => {
  const ext = hf as HeaderFooterWithVerbatim;
  delete ext.verbatimXml;
  delete ext.verbatimFingerprint;
};
