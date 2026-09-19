// Fixture for `folio-base64/no-hand-rolled-base64`. Encoding through the owner
// is the accepted spelling, and decoding is out of the rule's scope.

import { bytesToBase64, bytesToDataUrl } from "../../packages/core/src/utils/base64";

export const encode = (bytes: Uint8Array): string => bytesToBase64(bytes);

export const dataUrl = (bytes: Uint8Array): string => bytesToDataUrl(bytes, "image/png");

export const decode = (encoded: string): string => atob(encoded);
