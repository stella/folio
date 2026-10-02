import type { Document } from "../model/document";

const sourceReplayBrand = Symbol("SourceReplayToken");

/** Identity of a parser-registered immutable source; not an authentication boundary. */
export type SourceReplayToken = { readonly [sourceReplayBrand]: true };

const sourceReplayTokens = new WeakMap<Document, SourceReplayToken>();
const frozenGraphs = new WeakSet<object>();

const freezeGraph = (root: unknown): void => {
  const pending = [root];
  while (pending.length !== 0) {
    const value = pending.pop();
    if (
      value === null ||
      typeof value !== "object" ||
      frozenGraphs.has(value) ||
      ArrayBuffer.isView(value) ||
      value instanceof ArrayBuffer ||
      (typeof SharedArrayBuffer !== "undefined" && value instanceof SharedArrayBuffer) ||
      value instanceof Map ||
      value instanceof Set
    ) {
      continue;
    }
    frozenGraphs.add(value);
    for (const child of Object.values(value)) pending.push(child);
    Object.freeze(value);
  }
};

const freezeSourceReplayGraphs = (document: Document): void => {
  if (typeof process === "undefined" || process.env["NODE_ENV"] === "production") return;
  const { package: pkg } = document;
  freezeGraph(pkg.document.content);
  freezeGraph(pkg.document.background);
  freezeGraph(pkg.document.finalSectionProperties);
  freezeGraph(pkg.styles);
  freezeGraph(pkg.theme);
};

/** Called only by the parser's explicit tracked-source path. */
export const registerSourceReplayDocument = (document: Document): SourceReplayToken => {
  const existing = sourceReplayTokens.get(document);
  if (existing !== undefined) return existing;
  const token = Object.freeze({ [sourceReplayBrand]: true } as const satisfies SourceReplayToken);
  freezeSourceReplayGraphs(document);
  sourceReplayTokens.set(document, token);
  return token;
};

export const getSourceReplayToken = (document: Document): SourceReplayToken | undefined =>
  sourceReplayTokens.get(document);

/** Explicitly propagate source identity through an immutable document transformation. */
export const inheritSourceReplayToken = (target: Document, source: Document): void => {
  const token = sourceReplayTokens.get(source);
  if (token === undefined) return;
  freezeSourceReplayGraphs(target);
  sourceReplayTokens.set(target, token);
};
