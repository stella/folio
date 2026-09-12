/**
 * Paragraph attrs outside the typed Document projection must choose an owner.
 * `retained-in-memory` survives only through the private conversion channel;
 * `document-derived` is reconstructed from non-field Document metadata.
 */
type ParagraphProjectionOnlyAttrDescriptor = Readonly<{
  default: unknown;
  owner: "retained-in-memory" | "document-derived";
}>;

export const PARAGRAPH_PROJECTION_ONLY_ATTRS = Object.freeze({
  idStability: Object.freeze({
    default: undefined,
    owner: "retained-in-memory",
  }),
  _detachedWatermarkHost: Object.freeze({
    default: null,
    owner: "document-derived",
  }),
} as const satisfies Record<string, ParagraphProjectionOnlyAttrDescriptor>);

export type ParagraphProjectionOnlyAttrName = keyof typeof PARAGRAPH_PROJECTION_ONLY_ATTRS;

export const PARAGRAPH_ID_STABILITY_ATTR =
  "idStability" as const satisfies ParagraphProjectionOnlyAttrName;

export const POSITIONAL_PARAGRAPH_ID_STABILITY = "positional";
