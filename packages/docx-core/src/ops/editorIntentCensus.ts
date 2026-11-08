import { panic } from "better-result";

import type { BlockContent, Document } from "../model/document";
import { type ParagraphLocation, type storyParagraphs } from "./blocks";
import {
  countKeys,
  IDENTITY_SPACES,
  identityKeysIn,
  idKey,
  packageIdentityKeys,
  packageParagraphIds,
  packageReservedIdentityKeys,
  paragraphIdsIn,
  reservedIdentityKeysIn,
} from "./ids";
import type { OpStory } from "./types";

const sameFieldsExcept = (
  before: object,
  after: object,
  excluded: ReadonlySet<string>,
): boolean => {
  const beforeFields = Object.entries(before).filter(
    ([key]) => !excluded.has(key),
  );
  const afterFields = new Map(
    Object.entries(after).filter(([key]) => !excluded.has(key)),
  );
  return (
    beforeFields.length === afterFields.size &&
    beforeFields.every(
      ([key, value]) => afterFields.has(key) && afterFields.get(key) === value,
    )
  );
};

const PACKAGE_CONTENT_FIELDS = new Set(["document"]);
const DOCUMENT_CONTENT_FIELDS = new Set(["content", "sections"]);

type AdjustCountsOptions = {
  counts: Map<string, number>;
  keys: readonly string[];
  delta: number;
};

const adjustCounts = ({ counts, keys, delta }: AdjustCountsOptions): void => {
  for (const key of keys) {
    const count = (counts.get(key) ?? 0) + delta;
    if (count < 0)
      panic("An incremental identity census cannot have a negative count.");
    if (count === 0) counts.delete(key);
    else counts.set(key, count);
  }
};

const censusOf = (
  identities: Map<string, number>,
  paragraphs: Map<string, number>,
) => {
  const idsIn = (space: string) =>
    new Set(
      [...identities.keys()]
        .filter((key) => key.startsWith(`${space}:`))
        .map((key) => Number(key.slice(space.length + 1))),
    );
  return {
    revisions: idsIn(IDENTITY_SPACES.REVISION),
    controls: idsIn(IDENTITY_SPACES.CONTROL),
    paragraphs: new Set(paragraphs.keys()),
    stories: new Map<OpStory, ReturnType<typeof storyParagraphs>>(),
  };
};

type CensusVersion = {
  document: Document;
  identities: Map<string, number>;
  paragraphs: Map<string, number>;
  census: ReturnType<typeof censusOf>;
};

/** Cache immutable versions and update a main story without walking the blocks an edit left. */
export const createCensusReader = () => {
  const versions = new WeakMap<Document, CensusVersion>();
  const blockKeys = new WeakMap<
    BlockContent,
    { identities: string[]; paragraphs: string[] }
  >();
  let latest: CensusVersion | undefined;
  // A body block's ids are its own: a table's row and cell controls are
  // counted within that table, and nothing outside a block changes what it
  // holds, so a block counts the same wherever it stands.
  const keysOf = (block: BlockContent) => {
    let keys = blockKeys.get(block);
    if (keys === undefined) {
      keys = {
        identities: identityKeysIn(block).concat(reservedIdentityKeysIn(block)),
        paragraphs: paragraphIdsIn(block).map(idKey),
      };
      blockKeys.set(block, keys);
    }
    return keys;
  };
  return (
    document: Document,
  ): {
    revisions: Set<number>;
    controls: Set<number>;
    paragraphs: Set<string>;
    stories: Map<OpStory, ParagraphLocation[]>;
  } => {
    const cached = versions.get(document);
    if (cached !== undefined) {
      latest = cached;
      return cached.census;
    }
    const before = latest;
    const oldContent = before?.document.package.document.content;
    const content = document.package.document.content;
    let identities: Map<string, number>;
    let paragraphs: Map<string, number>;
    if (
      before !== undefined &&
      oldContent !== undefined &&
      sameFieldsExcept(
        before.document.package,
        document.package,
        PACKAGE_CONTENT_FIELDS,
      ) &&
      sameFieldsExcept(
        before.document.package.document,
        document.package.document,
        DOCUMENT_CONTENT_FIELDS,
      )
    ) {
      identities = new Map(before.identities);
      paragraphs = new Map(before.paragraphs);
      const changed = new Map<BlockContent, number>();
      for (const block of oldContent)
        changed.set(block, (changed.get(block) ?? 0) - 1);
      for (const block of content)
        changed.set(block, (changed.get(block) ?? 0) + 1);
      for (const [block, delta] of changed) {
        if (delta === 0) continue;
        const keys = keysOf(block);
        adjustCounts({ counts: identities, keys: keys.identities, delta });
        adjustCounts({ counts: paragraphs, keys: keys.paragraphs, delta });
      }
    } else {
      identities = countKeys(
        packageIdentityKeys(document.package).concat(
          packageReservedIdentityKeys(document.package),
        ),
      );
      paragraphs = countKeys(packageParagraphIds(document.package).map(idKey));
    }
    const version = {
      document,
      identities,
      paragraphs,
      census: censusOf(identities, paragraphs),
    };
    versions.set(document, version);
    latest = version;
    return version.census;
  };
};
