import type { CanonicalRefusalGround } from "./canonicalCapabilities";

export type CanonicalSectionPropertiesResult =
  | { status: "applied"; version: number }
  | { status: "refused"; ground: CanonicalRefusalGround; message: string };
