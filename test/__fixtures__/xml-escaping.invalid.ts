/** Deliberate violations: a writer that escapes XML by hand. */

export const escapeChain = (value: string): string =>
  value.replaceAll("&", "&amp;").replaceAll("<", "&lt;");

const ESCAPES: Record<string, string> = { "&": "&amp;" };

export const escapeTable = (value: string): string =>
  value.replace(/[&]/gu, (character) => ESCAPES[character] ?? character);
