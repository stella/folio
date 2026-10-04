/** Exact resource restoration preserves explicitly undefined own fields. */
import type { DocxPackage } from "../../src/model/document";

const PACKAGE = {
  document: { content: [] },
  styles: undefined,
  relationships: undefined,
  media: undefined,
} satisfies DocxPackage;

export type PackageResourcePresenceProof = typeof PACKAGE;
