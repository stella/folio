/**
 * The package's one bytes-to-base64 encoder.
 *
 * Every call site that spelled this itself got a different answer to the same
 * two questions, and one of them got the first question wrong:
 *
 * - `btoa(new TextDecoder("latin1").decode(bytes))` is not a binary string.
 *   The Encoding standard maps the `latin1` label to windows-1252, so bytes
 *   0x80-0x9F decode to characters outside Latin-1 and `btoa` throws
 *   `InvalidCharacterError` on them. Ordinary image bytes contain those values,
 *   so the browser path of that spelling fails on real input.
 * - `String.fromCharCode(...chunk)` is correct but spreads a chunk's worth of
 *   arguments per call, and the intermediate binary string is as large as the
 *   input.
 *
 * Neither question has to be asked per call site, so it is asked here once.
 * The runtime's own encoder answers it when the runtime has one, and the
 * portable encoder in the shared owner answers it when not; both produce canonical base64,
 * so the choice is invisible in the output.
 */

import { bytesToBase64 } from "@stll/utf8-boundaries";

export { bytesToBase64, encodeBase64Portable } from "@stll/utf8-boundaries";

/** Encode bytes as a `data:` URL with the given media type. */
export const bytesToDataUrl = (bytes: Uint8Array, mimeType: string): string =>
  `data:${mimeType};base64,${bytesToBase64(bytes)}`;
