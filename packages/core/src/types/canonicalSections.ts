import type { CANONICAL_GAP } from "./canonicalCapabilities";

export type CanonicalSectionPropertiesResult =
  | { status: "applied"; version: number }
  | { status: "refused"; gap: typeof CANONICAL_GAP.sectionProperties; message: string };
