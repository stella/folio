import type { DocxPackage } from "../model/document";
import {
  IDENTITY_SPACES,
  type IdentitySpace,
  packageIdentityKeyIndex,
  packageParagraphIdIndex,
  packageReservedKeyIndex,
  slotKey,
} from "./ids";

/** Which ids of a package are taken, asked one id at a time. */
export type PackageIdOccupancy = {
  /** Revision ids records carry or pending cut references reserve. */
  revisions: { has: (id: number) => boolean };
  /** Content-control ids, likewise. */
  controls: { has: (id: number) => boolean };
  /** Paragraph ids, by {@link idKey}. */
  paragraphs: { has: (key: string) => boolean };
};

/**
 * The ids a package holds, read from the package's kept key indexes: each is
 * derived from a package indexed before it along the paths an edit changed,
 * so a new version is answered without a census of the whole package.
 */
export const packageIdOccupancy = (pkg: DocxPackage): PackageIdOccupancy => {
  const identities = packageIdentityKeyIndex(pkg);
  const reserved = packageReservedKeyIndex(pkg);
  const paragraphs = packageParagraphIdIndex(pkg);
  const inSpace = (space: IdentitySpace) => ({
    has: (id: number) => {
      const key = slotKey({ space, id });
      return identities.has(key) || reserved.has(key);
    },
  });
  return {
    revisions: inSpace(IDENTITY_SPACES.REVISION),
    controls: inSpace(IDENTITY_SPACES.CONTROL),
    paragraphs: { has: (key) => paragraphs.has(key) },
  };
};
