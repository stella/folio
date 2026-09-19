// Fixture for `folio-base64/no-hand-rolled-base64`. The rule must flag every
// call, however the global is reached.

export const viaLatin1 = (bytes: Uint8Array): string =>
  btoa(new TextDecoder("latin1").decode(bytes));

export const viaChunks = (chunks: string[]): string => btoa(chunks.join(""));

export const viaGlobalThis = (binary: string): string => globalThis.btoa(binary);
