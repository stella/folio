/**
 * How TIFF images are decoded for display. Browsers cannot paint TIFF, so the
 * editor decodes it in a worker fetched the first time a document holds one,
 * and paints a placeholder when that fails; the document keeps the TIFF.
 *
 * A host that cannot load the worker (offline, a Content Security Policy
 * without `worker-src`, a single-file bundle) hands the decoder over:
 * `configureTiffDecoding({ decoder: await import("utif2") })`.
 *
 * @packageDocumentation
 */

export type { TiffDecoder } from "./utils/tiffDecode";
export {
  configureTiffDecoding,
  type TiffDecodingOptions,
} from "./utils/tiffConverter";
