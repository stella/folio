import type { DocxPackage } from "../model/document";
import {
  IDENTITY_SPACES,
  type IdentityKeySequence,
  type IdentitySpace,
  packageIdentityKeyIndex,
  packageParagraphIdIndex,
  packageReservedKeyIndex,
  paragraphIdSequence,
  slotKey,
} from "./ids";

/** Which ids of a package are taken, asked one id at a time. */
type IdOccupancy<Key> = {
  has: (key: Key) => boolean;
  occupiedPrefixEnd: (limit: number) => number;
};

export type PackageIdOccupancy = {
  /** Revision ids records carry or pending cut references reserve. */
  revisions: IdOccupancy<number>;
  /** Content-control ids, likewise. */
  controls: IdOccupancy<number>;
  /** Paragraph ids, by {@link idKey}. */
  paragraphs: IdOccupancy<string>;
};

const sequenceFor = (space: IdentitySpace): IdentityKeySequence => ({
  keyOf: (id) => slotKey({ space, id }),
  numberOf: (key) =>
    key.startsWith(`${space}:`)
      ? Number(key.slice(space.length + 1))
      : undefined,
});

const revisionSequence = sequenceFor(IDENTITY_SPACES.REVISION);
const controlSequence = sequenceFor(IDENTITY_SPACES.CONTROL);

/**
 * The ids a package holds, read from the package's kept key indexes: each is
 * derived from a package indexed before it along the paths an edit changed,
 * so a new version is answered without a census of the whole package.
 */
export const packageIdOccupancy = (pkg: DocxPackage): PackageIdOccupancy => {
  const identities = packageIdentityKeyIndex(pkg);
  const reserved = packageReservedKeyIndex(pkg);
  const paragraphs = packageParagraphIdIndex(pkg);
  const inSpace = (sequence: IdentityKeySequence) => ({
    has: (id: number) => {
      const key = sequence.keyOf(id);
      return identities.has(key) || reserved.has(key);
    },
    occupiedPrefixEnd: (limit: number) =>
      Math.max(
        identities.firstMissing(sequence, limit),
        reserved.firstMissing(sequence, limit),
      ),
  });
  return {
    revisions: inSpace(revisionSequence),
    controls: inSpace(controlSequence),
    paragraphs: {
      has: (key) => paragraphs.has(key),
      occupiedPrefixEnd: (limit) =>
        paragraphs.firstMissing(paragraphIdSequence, limit),
    },
  };
};
